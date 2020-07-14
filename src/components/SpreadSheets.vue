<template>
  <div class="home">
    <div class="toolbar">
      <!-- <el-upload
        class="upload-demo"
        action
        :auto-upload="false"
        :on-change="importSSExcel"
        :file-list="fileList"
      >
        <el-button size="small" type="primary">导入</el-button>
      </el-upload>-->
      <input type="file" class="el-button" @change="importExcel($event)" style />
      <el-button @click="exportExcel()">导出 Excel</el-button>
      <el-button icon="el-icon-download" type="primary" @click="download">下载</el-button>
      <el-button icon="el-icon-upload" type="primary" @click="upload">上传</el-button>
      <!-- <el-button icon="el-icon-lock" type="primary" @click="lockSpread(spread)">锁定表单</el-button>
      <el-button icon="el-icon-unlock" type="primary" @click="unlockSpread(spread)">解锁表单</el-button>
      <el-button type="primary" @click="getSelect(spread)">获取选择单元格行和列</el-button>-->

      <!-- <el-button-group style="margin-left:10px">
        <el-button type="primary" icon="el-icon-back" @click="hAlignLeft()"></el-button>
        <el-button type="primary" @click="hAlignCenter()">居中对齐</el-button>
        <el-button type="primary" icon="el-icon-right" @click="hAlignRight()"></el-button>
        <el-button type="primary" icon="el-icon-menu" @click="mergeCell()"></el-button>
      </el-button-group>-->
      <el-button type="primary" icon="el-icon-delete" @click="clearRange()">删除内容</el-button>

      <el-button-group style="margin-left:10px">
        <el-button type="primary" icon="el-icon-arrow-left" @click="undo()"></el-button>
        <el-button type="primary" icon="el-icon-arrow-right" @click="redo()"></el-button>
      </el-button-group>

      <el-button type="primary" @click="toJSON()">toJSON</el-button>
      <el-button type="primary" @click="fromJSON()">fromJSON</el-button>
    </div>
    <div class="spreadWrapper">
      <div ref="formulaBar" class="formulaBar" contenteditable="true" spellcheck="false"></div>
      <gc-spread-sheets class="spreadHost" @workbookInitialized="workbookInitialized($event)">
        <gc-worksheet :dataSource="dataTable"></gc-worksheet>
      </gc-spread-sheets>
      <div ref="statusBar" class="statusBar"></div>
    </div>
  </div>
</template>

<script lang="ts">
/* eslint-disable */
import Vue from "vue";
import { Notification } from "element-ui";
// import "@grapecity/spread-sheets/styles/gc.spread.sheets.excel2016colorful.css";
import "@grapecity/spread-sheets/styles/gc.spread.sheets.excel2013lightGray.css";
import "@grapecity/spread-sheets-vue";
import GC from "@grapecity/spread-sheets";
import ExcelIO from "@grapecity/spread-excelio";
import FileSaver from "file-saver";
// import '@grapecity/spread-sheets-charts';
import "@grapecity/spread-sheets-resources-zh";
import { Watch } from "vue-property-decorator";
GC.Spread.Common.CultureManager.culture("zh-cn");

export default class SpreadSheets extends Vue {
  spread: GC.Spread.Sheets.Workbook | null;
  spreadNS = GC.Spread.Sheets;
  fileList: object[] = [];
  dataTable:Array<object> = [{id:1,name:"张三",age:23},{id:2,name:"李四",age:23},{id:3,name:"王五",age:33}];

  constructor() {
    super();
    this.spread = null;
  }

  @Watch("$route", { immediate: true })
  private changeRouter(route: any) {
    // console.log(route)
    if (route.params.id) {
      this.initSpread(this.spread);
    }
  }
  workbookInitialized(spread: GC.Spread.Sheets.Workbook) {
    this.spread = spread;

    let formulaBar = new GC.Spread.Sheets.FormulaTextBox.FormulaTextBox(
      this.$refs.formulaBar as HTMLElement,
      {} as GC.Spread.Sheets.FormulaTextBox.IFormulaTextBoxOptions
    );
    formulaBar.workbook(this.spread);
    this.spread.focus();

    let statusBar = new GC.Spread.Sheets.StatusBar.StatusBar(
      this.$refs.statusBar as HTMLElement
    );
    statusBar.bind(spread);

    this.initSpread(spread);

    this.setMonthPicker(spread);

    this.registCommand(this.spread);
    this.registEvent(this.spread);
  }
  sendMessage(info: any) {
    let cellRef = GC.Spread.Sheets.CalcEngine.rangeToFormula(
      new GC.Spread.Sheets.Range(
        info.row,
        info.col,
        info.rowCount | 1,
        info.colCount | 1
      ),
      0,
      0,
      GC.Spread.Sheets.CalcEngine.RangeReferenceRelative.allRelative,
      false
    );
    let message = "单元格" + cellRef + "发生了变化！";

    Notification({ title: "同步", message: message, type: "info" });
  }
  registEvent(spread: GC.Spread.Sheets.Workbook) {
    let self = this;
    spread.bind(GC.Spread.Sheets.Events.ValueChanged, (s: any, e: any) => {
      self.sendMessage(e);
    });
    spread.bind(GC.Spread.Sheets.Events.RangeChanged, (s: any, e: any) => {
      self.sendMessage(e);
    });
  }
  lockSpread(spread: any) {
    if (spread) {
      console.log("lock");
      spread.options.newTabVisible = false;
      spread.options.allowContextMenu = false;
      let sheet = spread.getActiveSheet();
      sheet.options.isProtected = true;
    }
  }
  unlockSpread(spread: any) {
    if (spread) {
      console.log("lock");
      spread.options.newTabVisible = true;
      spread.options.allowContextMenu = true;
      let sheet = spread.getActiveSheet();
      sheet.options.isProtected = false;
    }
  }
  initSpread(spread: any) {
    console.log("init");
    let self = this;

    if (this.$route.params.id && self.spread) {
      console.log(this.$route.params.id);
      self.spread.options.newTabVisible = false;
      self.spread.options.allowContextMenu = false;
      let sheet = self.spread.getActiveSheet();
      sheet.options.isProtected = true;
    }
  }
  importExcel(event: any) {
    const file = event.target.files[0];
    console.log(file);
    let self = this;
    let excelIO = new ExcelIO.IO();
    excelIO.open(file, (spreadJSON: object) => {
      if (self.spread) {
        self.spread.fromJSON(spreadJSON);
      }
    });
    // this.initSpread(self.spread);
  }
  exportExcel() {
    let self = this;
    if (self.spread) {
      let spreadJSON = JSON.stringify(self.spread.toJSON());
      let excelIO = new ExcelIO.IO();
      excelIO.save(spreadJSON, (blob: Blob) => {
        FileSaver.saveAs(blob, "export.xlsx");
      });
    }
  }
  importSSExcel(file: any, fileList: any) {
    console.log(file.raw, fileList[0].raw);
    let self = this;
    let raw_file = fileList[0].raw;
    // let render = new FileReader();
    // render.readAsText(raw_file, "UTF-8");
    // render.onload = function(evt: any) {
    //   var fileString = evt.target.result as string;
    //   var jsonObj = JSON.parse(fileString);
    //   if (self.spread) {
    //     self.spread.fromJSON(jsonObj);
    //   }
    // };
    delete raw_file.uid;
    let excelIO = new ExcelIO.IO();
    excelIO.open(raw_file, (spreadJSON: object) => {
      if (self.spread) {
        self.spread.fromJSON(spreadJSON);
      }
    });
  }
  download() {}
  upload() {}
  getActualRange(range: any, maxRowCount: any, maxColCount: any) {
    var row = range.row < 0 ? 0 : range.row;
    var col = range.col < 0 ? 0 : range.col;
    var rowCount = range.rowCount < 0 ? maxRowCount : range.rowCount;
    var colCount = range.colCount < 0 ? maxColCount : range.colCount;

    return new this.spreadNS.Range(row, col, rowCount, colCount);
  }
  getSelect(spread: any) {
    if (spread) {
      var sheet = spread.getActiveSheet();
      var sels = sheet.getSelections();
      if (sels && sels.length > 0) {
        var sel = this.getActualRange(
          sels[0],
          sheet.getRowCount(),
          sheet.getColumnCount()
        );
        var comboBoxCellType = sheet.getCellType(sel.row, sel.col);
        if (!(comboBoxCellType instanceof this.spreadNS.CellTypes.ComboBox)) {
          // _getElementById("changeProperty").setAttribute("disabled", "disabled");
          alert(sel.row + "," + sel.col);
          return;
        } else {
          alert("Combo");
        }
      }
    }
  }

  setMonthPicker(spread: any) {
    if (this.spread) {
      // -------------------- Month Picker ---------------------
      let monthPickerStyle = new GC.Spread.Sheets.Style();
      let sheet = this.spread.getActiveSheet();
      monthPickerStyle.cellButtons = [
        {
          imageType: GC.Spread.Sheets.ButtonImageType.dropdown,
          command: "openMonthPicker",
          useButtonStyle: true
        }
      ];

      monthPickerStyle.dropDowns = [
        {
          type: GC.Spread.Sheets.DropDownType.monthPicker,
          option: {
            startYear: 2009,
            stopYear: 2019,
            height: 300
          }
        }
      ];
      sheet.setText(1, 5, "Month Picker");
      sheet.setStyle(2, 5, monthPickerStyle);
    }
  }

  hAlignLeft() {
    if (this.spread) {
      let sheet = this.spread.getActiveSheet();
      let range = sheet.getSelections()[0];
      sheet.suspendPaint();
      sheet
        .getRange(range.row, range.col, range.rowCount, range.colCount)
        .hAlign(GC.Spread.Sheets.HorizontalAlign.left);
      sheet.resumePaint();
    }
  }
  hAlignCenter() {
    if (this.spread) {
      let sheet = this.spread.getActiveSheet();
      let range = sheet.getSelections()[0];
      sheet.suspendPaint();
      sheet
        .getRange(range.row, range.col, range.rowCount, range.colCount)
        .hAlign(GC.Spread.Sheets.HorizontalAlign.center);
      sheet.resumePaint();
    }
  }
  hAlignRight() {
    if (this.spread) {
      let sheet = this.spread.getActiveSheet();
      let range = sheet.getSelections()[0];
      sheet.suspendPaint();
      sheet
        .getRange(range.row, range.col, range.rowCount, range.colCount)
        .hAlign(GC.Spread.Sheets.HorizontalAlign.right);
      sheet.resumePaint();
    }
  }
  mergeCell() {
    if (this.spread) {
      let commandManager = this.spread.commandManager();
      let sheet = this.spread.getActiveSheet();
      commandManager.execute({
        cmd: "mergeCellCommand",
        sheetName: sheet.name(),
        selections: sheet.getSelections()
      });
    }
  }
  clearRange() {
    if (this.spread) {
      let commandManager = this.spread.commandManager();
      let sheet = this.spread.getActiveSheet();
      commandManager.execute({
        cmd: "clearRangeCommand",
        sheetName: sheet.name(),
        selections: sheet.getSelections()
      });
    }
  }
  registCommand(spread: GC.Spread.Sheets.Workbook) {
    let commandManager = spread.commandManager();

    let mergeCellCommand = {
      canUndo: true,
      execute: function(
        spread: GC.Spread.Sheets.Workbook,
        options: any,
        isUndo: boolean
      ) {
        var Commands = GC.Spread.Sheets.Commands;
        if (isUndo) {
          Commands.undoTransaction(spread, options);
          return true;
        } else {
          Commands.startTransaction(spread, options);
          spread.suspendPaint();
          var selections = options.selections;
          var sheet = spread.getSheetFromName(options.sheetName);
          selections.forEach(function(sel: GC.Spread.Sheets.Range) {
            if (sel.rowCount > 1 || sel.colCount > 1) {
              sheet.addSpan(sel.row, sel.col, sel.rowCount, sel.colCount);
            }
          });
          spread.resumePaint();
          Commands.endTransaction(spread, options);
          return true;
        }
      }
    };
    commandManager.register("mergeCellCommand", mergeCellCommand);

    let clearRangeCommand = {
      canUndo: true,
      execute: function(
        spread: GC.Spread.Sheets.Workbook,
        options: any,
        isUndo: boolean
      ) {
        var Commands = GC.Spread.Sheets.Commands;
        if (isUndo) {
          Commands.undoTransaction(spread, options);
          return true;
        } else {
          Commands.startTransaction(spread, options);
          spread.suspendPaint();
          var selections = options.selections;
          var sheet = spread.getSheetFromName(options.sheetName);
          selections.forEach(function(sel: GC.Spread.Sheets.Range) {
            sheet.clear(
              sel.row,
              sel.col,
              sel.rowCount,
              sel.colCount,
              GC.Spread.Sheets.SheetArea.viewport,
              GC.Spread.Sheets.StorageType.data
            );
          });
          spread.resumePaint();
          Commands.endTransaction(spread, options);
          return true;
        }
      }
    };
    commandManager.register("clearRangeCommand", clearRangeCommand);
  }

  undo() {
    if (this.spread) {
      let undoManager = this.spread.undoManager();
      undoManager.undo();
    }
  }
  redo() {
    if (this.spread) {
      let undoManager = this.spread.undoManager();
      undoManager.redo();
    }
  }
  toJSON() {
    var serializationOption = {
      ignoreFormula: true,
      ignoreStyle: true,
      rowHeadersAsFrozenColumns: true,
      columnHeadersAsFrozenRows: true
    };
    if (this.spread) {
      var jsonStr = JSON.stringify(this.spread.toJSON(serializationOption));
      console.log(jsonStr);
    }
  }
  fromJSON() {
    var jsonOptions = {
      ignoreFormula: true,
      ignoreStyle: true,
      frozenColumnsAsRowHeaders: true,
      frozenRowsAsColumnHeaders: true,
      doNotRecalculateAfterLoad: true
    };
    let jsonStr = {
      version: "13.1.2",
      customList: [],
      sheets: {
        Sheet1: {
          name: "Sheet1",
          isSelected: true,
          rowCount: 201,
          columnCount: 21,
          activeRow: 4,
          activeCol: 6,
          data: {
            dataTable: {
              "1": {
                "1": { value: "id" },
                "2": { value: "name" },
                "3": { value: "gender" },
                "4": { value: "age" }
              },
              "2": {
                "1": { value: 1 },
                "2": { value: "张三" },
                "3": { value: "男" },
                "4": { value: 21 }
              },
              "3": {
                "1": { value: 2 },
                "2": { value: "李四" },
                "3": { value: "女" },
                "4": { value: 21 }
              },
              "4": {
                "1": { value: 3 },
                "2": { value: "王五" },
                "3": { value: "男" },
                "4": { value: 23 }
              },
              "5": {
                "1": { value: 4 },
                "2": { value: "赵六" },
                "3": { value: "女" },
                "4": { value: 18 }
              }
            }
          },
          rowHeaderData: {},
          colHeaderData: {},
          rows: [
            null,
            null,
            null,
            { size: 20, visible: true },
            { size: 24, visible: true },
            { size: 20, visible: true }
          ],
          columns: [{ size: 40 }],
          leftCellIndex: 1,
          topCellIndex: 1,
          selections: {
            "0": { row: 4, rowCount: 1, col: 6, colCount: 1 },
            length: 1
          },
          outlineColumnOptions: {},
          index: 0
        }
      }
    };
    if (this.spread) {
      this.spread.fromJSON(jsonStr, jsonOptions);
    }
  }
}
</script>

<style>
.home,
.spreadWrapper {
  height: calc(100% - 45px);
}
.spreadHost {
  height: calc(100% - 75px);
  width: 100%;
}
.toolbar {
  padding-bottom: 5px;
  display: flex;
  align-items: center;
}
.formulaBar {
  height: 43px;
  width: calc(100% - 3px);
  margin-bottom: 2px;
  border: 1px solid #808080;
  background: white;
}
.statusBar {
  height: 25px;
  width: 100%;
}
</style>
