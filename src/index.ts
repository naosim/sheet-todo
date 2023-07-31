type Config = {
  sheetName: string,
  backupSheetName: string,
  googleTaskListId: string, 
}
const config: Config = {
  sheetName: 'シート1',
  backupSheetName: 'タスクバックアップ',
  googleTaskListId: null, // ★設定してください
}


declare var SpreadsheetApp: any;
declare var Utilities: {formatDate: (date: Date, locale: string, format: string) => string}
declare var Tasks: {Tasks: any};

// いい感じに同期します
// 新規タスクの追加
// 変更点の更新
// GoogleTasksの読み込み
function update() {
  new MainService(config).update();
}

function onOpen() {
  const customMenu = SpreadsheetApp.getUi()
  customMenu.createMenu('タスク')
    .addItem('更新', 'update') 
    .addToUi()
}

// スプレッドシートをセットアップします。
function setupSheet() {
  new MainService(config).setupSheet();
}

class Sheet {
  private sheet: any;
  sheetName: string;
  constructor(sheetName: string) {
    this.sheetName = sheetName;
  }
  getSheet(): any {
    if(!this.sheet) {
      var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
      var a = spreadSheet.getSheetByName(this.sheetName);
      if(a) {
        this.sheet = a;
      } else {
        this.sheet = spreadSheet.insertSheet();
        this.sheet.setName(this.sheetName);
      }
    }
    return this.sheet;

  }
  clear() {
    this.getSheet().clear();
  }
  clearContents() {
    this.getSheet().clearContents();
  }
  setValues(values: any[][], options?: {startRow: number, startColumn: number, rowLength: number, columnLength: number}) {
    if(!options) {
      options = {startRow: 1, startColumn: 1, rowLength: values.length, columnLength: values[0].length}
    }
    this.getSheet().getRange(options.startRow, options.startColumn, options.rowLength, options.columnLength).setValues(values);
  }
  getValues(): any[][] {
    return this.getSheet().getDataRange().getValues() as any[][]
  }
}

class MainService {
  private googleTasksRepository: GoogleTasksRepository;
  private taskSheetRepository: TaskSheetRepository;
  constructor(config: Config) {
    // DI
    this.googleTasksRepository = new GoogleTasksRepository(config.sheetName);
    this.taskSheetRepository = new TaskSheetRepository(new Sheet(config.sheetName), new Sheet(config.backupSheetName));
  }

  static dateToText(dateOrText: string | Date): string {
    if(!dateOrText) {
      return dateOrText as string;
    }
    if(toString.call(dateOrText).indexOf('Date') == -1) {
      return dateOrText as string;
    }
    return (dateOrText as Date).toLocaleDateString();

  }

  static convertGoogleTaskToSheetTask(googleTask: GoogleTask): SheetTask {
    const short = [googleTask.title, MainService.dateToText(googleTask.dates.due), googleTask.notes].filter(v => v).map(v => v.trim()).join('\n').slice(0, 140);
    const result = {
      id: googleTask.id,
      title: googleTask.title,
      notes: googleTask.notes,
      due: googleTask.dates.due,
      completed: googleTask.dates.completed,
      updated: googleTask.dates.updated,
      short
    }

    return result;
  }


  update() {
    // シート上の変更をGoogleTaskに反映する
    const diff = this.taskSheetRepository.getDiffTasks();
    const newTaskItems = diff.filter(v => v.status == 'new');
    const updateTaskItems = diff.filter(v => v.status == 'update');
    newTaskItems.forEach(v => {
      this.googleTasksRepository.insert(v.data);
    })
    updateTaskItems.forEach(v => {
      this.googleTasksRepository.update(v.id, v.data);
    })

    // googleからタスクを取得
    const googleTasks = this.googleTasksRepository.getTasks();
    // sheet用に変換
    const sheetTasks = googleTasks.map((v:GoogleTask) => MainService.convertGoogleTaskToSheetTask(v));
    // sheetへ保存
    this.taskSheetRepository.updateSheet(sheetTasks);
  }

  setupSheet() {
    this.taskSheetRepository.setupSheet();
  }
}

type JudgeResult = {
  status: 'none' | 'new' | 'update', id?: string, data: {title?: string, notes?: string, due?: Date, completed?: Date}
}

class DiffJudge {
  judge(s: SheetTask, b: SheetTask): JudgeResult {
    const eqDate = (a: Date, b: Date) => {
      if(!a && !b) {
        return true;
      }
      if(a && !b) {
        return false;
      }
      if(!a && b) {
        return false;
      }
      return a.getTime() == b.getTime();
    };

    const result: JudgeResult = {status: 'none', id: s.id, data: {title: undefined, notes: undefined, due: undefined, completed: undefined}};
    if(!s.id) {
      return {status: 'new', data: s};
    }
    if(s.title != b.title) {
      result.status = 'update';
      result.data.title = s.title;
    }
    if(s.notes != b.notes) {
      result.status = 'update';
      result.data.notes = s.notes;
    }
    if(!eqDate(s.due, b.due)) {
      result.status = 'update';
      result.data.due = s.due;
    }
    if(!eqDate(s.completed, b.completed)) {
      result.status = 'update';
      result.data.completed = s.completed;
    }
    return result;
  }
}

type SheetTask = {
  'short': string,
  'completed': Date,
  'title': string,
  'notes': string,
  'due': Date,
  'id': string,
  'updated': Date
}

class TaskSheetRepository {

  _columns: string[];
  sheet: Sheet;
  backupSheet: Sheet;
  constructor(sheet: Sheet, backupSheet: Sheet) {
    this.sheet = sheet;
    this.backupSheet = backupSheet;
    this._columns = [
      'short',
      'completed',
      'title',
      'notes',
      'due',
      'id',
      'updated',
    ]
  }

  setupSheet() {
    this.sheet.clear();
    this.sheet.setValues([this._columns]);
    this.backupSheet.setValues([this._columns]);
  }

  updateSheet(sheetTasks: SheetTask[]) {
    const table = [
      this._columns,
      ...sheetTasks.map(t => this._columns.map(c => t[c]))
    ];

    const updateSheet = (sheet: Sheet) => {
      sheet.clearContents();
      sheet.setValues(table);
    }

    updateSheet(this.backupSheet);
    updateSheet(this.sheet);
    
  }


  getDiffTasks() {
    const aryToTask = (ary: any[]) => {
      return ary.reduce((memo: any, v, i) => {
        memo[this._columns[i]] = v;
        return memo;
      }, {}) as SheetTask
    }
    const sheetValues = this.sheet.getValues().slice(1).map(ary => aryToTask(ary));
    const backupValues = this.backupSheet.getValues().slice(1).map(ary => aryToTask(ary));

    return sheetValues.map((s, i) => {
      const b = backupValues[i];
      return new DiffJudge().judge(s, b);
    }).filter(v => v.status != 'none');
  
  }
}

type NewTask = {
  title?: string,
  notes?: string,
  due?: Date,
}
class GoogleTasksRepository {
  taskListId: string;
  constructor(taskListId: string) {
    this.taskListId = taskListId;
    if(!this.taskListId) {
      throw new Error('GoogleTasksRepository.taskListIdを設定してください');
    }
  }
  insert(task: NewTask) {
    /*
    {
      title: string,
      notes: string,
      due: Date,
    }
    */

    const input: {title: string, notes: string, due?: string} = {
      title: task.title,
      notes: task.notes
    };
    if(task.due) {
      input.due = Utilities.formatDate(task.due, "Asia/Tokyo", "yyyy-MM-dd") + "T00:00:00.000Z"
    }

    Tasks.Tasks.insert(input, this.taskListId);

  }

  update(id: string, task: {title?: string, notes?: string, due?: Date, completed?: Date}) {
    var t: any = {};
    if(task.title) {
      t.title = task.title
    }
    if(task.notes) {
      t.notes = task.notes
    }
    if(task.due) {
      t.due = Utilities.formatDate(task.due, "Asia/Tokyo", "yyyy-MM-dd") + "T00:00:00.000Z"
    }
    if(task.completed) {
      t.completed = Utilities.formatDate(task.completed, "Asia/Tokyo", "yyyy-MM-dd") + "T00:00:00.000Z"
    }

    Tasks.Tasks.patch(t, this.taskListId, id)
  }
  
  getTasks(): GoogleTask[] {
    return Tasks.Tasks.list(this.taskListId, {
      showCompleted: true,
      showHidden: true
    }).items.map((v: any) => {
      v.dates = {};
      // 日付をDate型に変える
      const keys = ['updated', 'completed', 'due'];
      keys.forEach(k => {
        if(v[k]) {
          v.dates[k] = new Date(v[k]);
        }
      })
      return v;
    });
  }
}

type GoogleTask = {
  "kind": string,
  "id": string,
  "etag": string,
  "title": string,
  "updated": string,
  "selfLink": string,
  "parent": string,
  "position": string,
  "notes": string,
  "status": string,
  "due": string,
  "completed": string,
  "deleted": boolean,
  "hidden": boolean,
  "links": [
    {
      "type": string,
      "description": string,
      "link": string
    }
  ],
  dates: {
    updated: Date, completed: Date, due: Date
  }
}