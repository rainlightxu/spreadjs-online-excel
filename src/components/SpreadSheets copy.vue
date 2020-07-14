
<template>
      <div class="home">
        <div class="toolbar">
          <!-- <el-upload
            class="upload-demo"
            style="width:200px;float:left;"
            action
            :auto-upload="false"
            :on-change="importSSExcel"
            ref="upload"
          >
            <el-button size="small" type="primary">导入</el-button>
          </el-upload>-->
          <!-- <el-input v-model="listQuery.Id" placeholder="CalendarID" class="filter-item"></el-input>
          <el-input v-model="listQuery.SKUCode" placeholder="渠道分配" class="filter-item"></el-input>
          <el-select v-model="listQuery.Lang" placeholder="客户组" class="filter-item">
            <el-option value="0">RTM</el-option>
            <el-option value="1">WTS</el-option>
          </el-select>
          <el-input v-model="listQuery.SKUDesc" placeholder="区域" class="filter-item"></el-input>
          <div style="display:inline-block;float:right;margin:0 10px 10px 0;">
            <span>活动月份:</span>
            <el-date-picker
              v-model="listQuery.StartDate"
              type="date"
              placeholder="开始日期"
              format="yyyy 年 MM 月 dd 日"
              value-format="yyyy-MM-dd"
              @change="handleFilter"
            ></el-date-picker>
            <span>-</span>
            <el-date-picker
              v-model="listQuery.EndDate"
              type="date"
              placeholder="结束日期"
              format="yyyy 年 MM 月 dd 日"
              value-format="yyyy-MM-dd"
              @change="handleFilter"
            ></el-date-picker>
          </div>

          <el-button
            icon="el-icon-search"
            type="primary"
            @click="handleFilter"
            class="filter-item"
          >搜索</el-button>-->
          <!-- <input
            type="file"
            class="el-button"
            @change="importExcel($event)"
            style="height: 45px;margin-right:10px;"
          />-->
          <div style="display:inline-block;float:right;">
            <el-button icon="el-icon-download" type="primary" size @click="exportExcel">导出到Excel</el-button>
            <el-button icon="el-icon-upload" type="primary" @click="upload">上 传</el-button>
            <el-button type="primary" @click="fromJSON()">fromJSON</el-button>
            <el-button type="primary" @click="bindDataSource()">绑定表单</el-button>
          </div>
          <!-- <el-button icon="el-icon-lock" type="primary" @click="lockSpread(spread)">锁定表单</el-button>
          <el-button icon="el-icon-unlock" type="primary" @click="unlockSpread(spread)">解锁表单</el-button>-->
        </div>

        <div class="spreadWrapper">
          <!-- <div ref="formulaBar" class="formulaBar" contenteditable="true" spellcheck="false"></div> -->
          <gc-spread-sheets
            class="spreadHost"
            v-loading="listLoading"
            @workbookInitialized="workbookInitialized($event)"
          >
            <gc-worksheet></gc-worksheet>
          </gc-spread-sheets>
          <!-- <div ref="statusBar" class="statusBar"></div>  :dataSource="dataTable"-->
        </div>
        <span id="tableColumnsContainer"></span>
      </div>
</template> 
<script>
/* eslint-disable */
import Vue from "vue";

import { Message } from "element-ui";

// import "@grapecity/spread-sheets/styles/gc.spread.sheets.excel2016colorful.css";
// import "@grapecity/spread-sheets/styles/gc.spread.sheets.excel2013darkGray.css";
import "@grapecity/spread-sheets/styles/gc.spread.sheets.excel2013lightGray.css";
import "@grapecity/spread-sheets-vue";
import GC from "@grapecity/spread-sheets";
import ExcelIO from "@grapecity/spread-excelio";
import FileSaver from "file-saver";
import "@grapecity/spread-sheets-resources-zh";
import PC from "./PC.json";
import HH from "./HH.json";
import ssJSON from "./ssJSON.json";
import dataSource from "./datasource.js";
import datasource from "./datasource.js";
GC.Spread.Common.CultureManager.culture("zh-cn");
GC.Spread.Sheets.LicenseKey =
  "exceldemo.softorg.com,199297176362428#B025mWQyUnb8YnV5clMlVUTplXZjNmbWd6UzJXWD9URJFWWxZTUplWeX3SerUXRLJVN9djSwhmcohWVtRkMvtycxUGNaBVa6tyMF5UeNJlY7ZEd0lVWN3SeGl4LLNXSwxmVVh6ZKR6NwkGcUR5Z9VEeUd6VhpmMspVSQRXbT54Smt4NLJVMvEXaXJVTi3mZJ9mQzgkaXpVSzV7NkhnUmpXO4MkWp9kaaNWZk56an5GZjFlcaBHdzU4cqVjRGNjQEhEUXdUc0ZGbxskdQdXcspXOQRDSBF4NvIXTiZEZ7JkY9F7VsN6bp36Z8klMxETQuVXaysWMSNDbRFzUzgFWTVXdvsmUiojITJCLiUkMFNUO5QUMiojIIJCL6ITOzYjMyITO0IicfJye&Qf35VfiU5TzYjI0IyQiwiIzEjL6ByUKBCZhVmcwNlI0IiTis7W0ICZyBlIsIyMxITNwEDIyEjNwAjMwIjI0ICdyNkIsIiMxcDMwIDMyIiOiAHeFJCLi46bj9yZy3Gdm36cu2WblRGblNGelJiOiMXbEJCLig1jlzahlDZmpnInm/KnmDoim/agmH0vkDJllb0pnfbtmrIukLiOiEmTDJCLigjM4IjNzYzNxcTOykTOxIiOiQWSiwSflNHbhZmOiI7ckJye0ICbuFkI1pjIEJCLi4TPBR4K6l6VzJjdNJWZk3maip6UMlEaFJ5VY5EMHdGTwk6ZSlTWlJWaMFzdwkETlJzcSllWS9GbQRkNadWY7cXe42GeSdTM5pmclV4L8o5QzkzzDr2";
export default {
  data() {
    return {
      PC: [],
      HH: [],
      fileList: [],
      autoGenerateColumns: true,
      width: 300,
      visible: true,
      resizable: true,
      formatter: "",
      spread: null,
      spreadNS: GC.Spread.Sheets,
      listQuery: {
        Id: "",
        SKUCode: "",
        SKUDesc: "",
        SapCode: "",
        Lang: "",
        startDate: "",
        endDate: ""
      },
      listLoading: true,
      selectedArray: null,
      selectedComboBoxArray: null,
      customerGroup: [],
      storePoolGroup: [],
      customerGroupCombo: null,
      storePoolGroupCombo: null,
      activityMonth: null,
      posterDate: null,
      startDate: null,
      endDate: null,
      schedule: null,
      scheduleGroup: []
    };
  },
  created() {
    // this.getList();
    // this.initList();
  },
  computed: {
    // dataTable() {
    // let dataTable = [];
    // for (let i = 0; i < 10; i++) {
    //   dataTable.push({
    //     id: i + 1,
    //     name: "路人" + i + 1,
    //     gender: Math.random < 0.5 ? "男" : "女",
    //     age: i + 10
    //   });
    // }
    // return dataTable;
    // }
  },
  methods: {
    // getList() {
    //   this.$axios
    //     .get(
    //       `/api/COA/GetCoaList?SkuCode=&SkuDesc=&SapCode=&BackCode=&StartDate=&EndDate=&Channel=&COALanguage=${"中文"}&page=1&limit=10&sort=-1`
    //     )
    //     .then(res => {
    //       console.log(res);
    //       this.dataTable = res.data.ReturnObject.ViewList;
    //       var sheet = this.spread.getActiveSheet();
    //       if (this.dataTable.length && this.dataTable.length > 0) {
    //         sheet.setDataSource(this.dataTable);
    //         this.initSpread(this.spread);
    //         this.hAlignCenter();
    //       }
    //     });
    // },
    initList() {
      this.PC = PC;
      this.HH = HH;
      this.customerGroup = [
        { text: "RTM", value: "0" },
        { text: "WTS", value: "1" },
        { text: "WALMART", value: "2" },
        { text: "永辉", value: "3" },
        { text: "carrefour", value: "4" }
      ];
      this.storePoolGroup = [
        { text: "N", value: "0" },
        { text: "Y", value: "1" }
      ];
      this.scheduleGroup = [
        {
          text: "02.14-02.16",
          value: "02.14-02.16"
        },
        {
          text: "02.21-02.23",
          value: "02.21-02.23"
        },
        {
          text: "02.28-03.01",
          value: "02.28-03.01"
        },
        {
          text: "03.06-03.08",
          value: "03.06-03.08"
        }
      ];
    },
    handleFilter() {},
    workbookInitialized(spread) {
      this.spread = spread;
      // let formulaBar = new GC.Spread.Sheets.FormulaTextBox.FormulaTextBox(
      //   this.$refs.formulaBar,
      //   {}
      // );
      // formulaBar.workbook(this.spread);
      // this.spread.focus();

      // let statusBar = new GC.Spread.Sheets.StatusBar.StatusBar(
      //   this.$refs.statusBar
      // );
      // statusBar.bind(spread);
      let sheet_HH = this.spread.getSheet(0);
      sheet_HH.name("HH");
      let sheet_PC = new GC.Spread.Sheets.Worksheet("PC");
      this.spread.addSheet(1, sheet_PC);
      //初始化spread
      this.initSpread(this.spread);
      //注册事件
      this.registEvent(this.spread);
      //注册命令
      this.registCommand(this.spread);
    },
    initSpread(spread) {
      this.listLoading = false;
      let self = this;
      //加载数据源
      this.initList();
      //获取sheet
      let sheet = spread.getSheet(0);
      let sheet_PC = spread.getSheet(1);
      //渲染数据
      if (this.HH.length && this.HH.length > 0) {
        sheet.setDataSource(this.HH);
      }
      if (this.PC.length && this.PC.length > 0) {
        sheet_PC.setDataSource(this.PC);
      }
      //行数和列数
      let rowCount = sheet.getRowCount();
      let colCount = sheet.getColumnCount();
      //设置不能新增表单
      spread.options.newTabVisible = false;
      // 将前18列的右键菜单屏蔽
      this.disableContextMenu(spread);
      //右键菜单选项删除
      this.removeMenuDataItems(spread);
      //隐藏sheet名称
      //   spread.options.tabStripVisible = false;

      //设置激活单元格
      sheet.setActiveCell(0, 0);
      sheet_PC.setActiveCell(0, 0);
      //设置filter
      var filter = new this.spreadNS.Filter.HideRowFilter(
        new this.spreadNS.Range(-1, 0, -1, colCount)
      );
      var filter_PC = new this.spreadNS.Filter.HideRowFilter(
        new this.spreadNS.Range(-1, 0, -1, colCount)
      );
      sheet.rowFilter(filter);
      sheet_PC.rowFilter(filter_PC);
      //设置列宽
      this.initColumnWidth(sheet, sheet_PC);
      //设置列头高
      var colHeader = GC.Spread.Sheets.SheetArea.colHeader;
      sheet.setRowHeight(0, 50, colHeader);
      sheet.setRowHeight(1, 30, colHeader);
      sheet_PC.setRowHeight(0, 50, colHeader);
      sheet_PC.setRowHeight(1, 30, colHeader);

      //设置下拉框
      this.customerGroupCombo = new GC.Spread.Sheets.CellTypes.ComboBox();
      this.storePoolGroupCombo = new GC.Spread.Sheets.CellTypes.ComboBox();
      var combo2 = new GC.Spread.Sheets.CellTypes.ComboBox();
      this.customerGroupCombo.items(this.customerGroup);
      this.storePoolGroupCombo.items(this.storePoolGroup);
      combo2.items([
        { text: "中文", value: "0" },
        { text: "英文", value: "1" }
      ]);

      //设置日期选择框
      // -------------------- Date Time Picker : showTime False ---------------------
      this.startDate = new GC.Spread.Sheets.Style();
      this.startDate.cellButtons = [
        {
          imageType: GC.Spread.Sheets.ButtonImageType.dropdown,
          command: "openDateTimePicker",
          useButtonStyle: true
        }
      ];
      this.startDate.dropDowns = [
        {
          type: GC.Spread.Sheets.DropDownType.dateTimePicker,
          option: {
            showTime: false
          }
        }
      ];
      this.endDate = new GC.Spread.Sheets.Style();
      this.endDate.cellButtons = [
        {
          imageType: GC.Spread.Sheets.ButtonImageType.dropdown,
          command: "openDateTimePicker",
          useButtonStyle: true
        }
      ];
      this.endDate.dropDowns = [
        {
          type: GC.Spread.Sheets.DropDownType.dateTimePicker,
          option: {
            showTime: false
          }
        }
      ];

      // -------------------- Vertical text list ---------------------
      this.schedule = new GC.Spread.Sheets.Style();
      this.schedule.cellButtons = [
        {
          imageType: GC.Spread.Sheets.ButtonImageType.dropdown,
          command: "openList",
          useButtonStyle: true
        }
      ];
      this.schedule.dropDowns = [
        {
          type: GC.Spread.Sheets.DropDownType.list,
          option: {
            multiSelect: true,
            items: this.scheduleGroup
          }
        }
      ];
      //设置下拉框和日期选择框,列表
      this.setComboBox(spread);
      this.setDateStyle(spread);
      this.setListStyle(spread);
      // sheet.setCellType(i, 8, combo2, GC.Spread.Sheets.SheetArea.viewport);
      // sheet.setStyle(i, 4, this.activityMonth);

      // sheet.setFormatter(i, 13, "0%");
      // sheet.setFormatter(i, 14, "￥#,##0.00");
      // sheet.getCell(i,4).formatter(new  GC.Spread.Formatter.GeneralFormatter("yyyy/MM/dd", "en-us"))
      //设置多级表头并上色
      this.setMultiColHeader(spread);
      //设置居中显示
      this.hAlignCenter();
      this.vAlignCenter();

      this.listLoading = false;
    },
    registEvent(spread) {
      let sheet = spread.getActiveSheet();
      let self = this;

      spread.bind(GC.Spread.Sheets.Events.ValueChanged, (s, args) => {
        console.log(args, "--valueChanged");
      });
      //复制粘贴的时候的确会触发rangeChanged事件
      spread.bind(GC.Spread.Sheets.Events.RangeChanged, (s, e) => {
        console.log(e, "--rangeChanged");
        //遍历每一个cell
        // e.changedCells.forEach(item => {
        //   if (item.col === 2) {
        //     //客户组
        //     let cellValue = sheet.getValue(item.row, item.col);
        //     console.log(cellValue);
        //     let existFlag = this.customerGroup.some(item => {
        //       return item.text == cellValue;
        //     });
        //     if (!existFlag) {
        //       console.log("不存在");
        //       // sheet.setValue(item.row,item.col,"")
        //       sheet
        //         .getCell(item.row, item.col, this.spreadNS.SheetArea.viewport)
        //         .cellType(this.customerGroupCombo)
        //         .value("");
        //     } else {
        //       console.log("存在");
        //       sheet
        //         .getCell(item.row, item.col, this.spreadNS.SheetArea.viewport)
        //         .cellType(this.customerGroupCombo)
        //         .value(cellValue);
        //     }
        //   }
        //   if (item.col === 4) {
        //     let cellValue = sheet.getValue(item.row, item.col);
        //     console.log(cellValue);
        //     // let reg = /\\\\d{4}(\-|\/|.)\\\\d{1,2}\1\\\\d{1,2}/
        //     // let reg = /\\\\d{4}\\\\d{1,2}\1\\\\d{1,2}/
        //     let reg = /^\d{4}\d{2}$/;
        //     let validFlag = reg.test(cellValue);
        //     if (!validFlag) {
        //       console.log("格式不正确");
        //       sheet.setStyle(item.row, item.col, this.activityMonth);
        //       sheet.setValue(item.row, item.col, "");
        //     } else {
        //       console.log("格式正确");
        //       sheet.setStyle(item.row, item.col, this.activityMonth);
        //       sheet.setValue(item.row, item.col, cellValue);
        //     }
        //   }
        //   if (item.col === 5) {
        //     let cellValue = sheet.getValue(item.row, item.col);
        //     console.log(cellValue);
        //     // let reg = /\\\\d{4}(\-|\/|.)\\\\d{1,2}\1\\\\d{1,2}/
        //     // let reg = /\\\\d{4}\\\\d{1,2}\1\\\\d{1,2}/
        //     let reg = /^\d{4}\d{2}\d{2}$/;
        //     let validFlag = reg.test(cellValue);
        //     if (!validFlag) {
        //       console.log("格式不正确");
        //       sheet.setStyle(item.row, item.col, this.posterDate);
        //       sheet.setValue(item.row, item.col, "");
        //     } else {
        //       console.log("格式正确");
        //       sheet.setStyle(item.row, item.col, this.posterDate);
        //       sheet.setValue(item.row, item.col, cellValue);
        //     }
        //   }
        //   if (item.col === 6) {
        //     let cellValue = sheet.getValue(item.row, item.col);
        //     console.log(cellValue);
        //     // let reg = /\\\\d{4}(\-|\/|.)\\\\d{1,2}\1\\\\d{1,2}/
        //     // let reg = /\\\\d{4}\\\\d{1,2}\1\\\\d{1,2}/
        //     let reg = /^\d{4}\d{2}\d{2}$/;
        //     let validFlag = reg.test(cellValue);
        //     if (!validFlag) {
        //       console.log("格式不正确");
        //       sheet.setStyle(item.row, item.col, this.posterDate);
        //       sheet.setValue(item.row, item.col, "");
        //     } else {
        //       console.log("格式正确");
        //       sheet.setStyle(item.row, item.col, this.posterDate);
        //       sheet.setValue(item.row, item.col, cellValue);
        //     }
        //   }
        // });
        //行高
        this.vAlignCenter();
        this.hAlignCenter();
      });

      spread.bind(GC.Spread.Sheets.Events.SelectionChanged, (s, e) => {
        // console.log(e, "--selectionChanged");
        e.newSelections.forEach(item => {
          let cellType = sheet.getCellType(item.row, item.col);
          if (
            cellType &&
            cellType instanceof GC.Spread.Sheets.CellTypes.ComboBox
          ) {
            // console.log("7")
          }
        });
      });
      sheet.bind(GC.Spread.Sheets.Events.CellClick, function(sender, args) {
        if (args.sheetArea === GC.Spread.Sheets.SheetArea.colHeader) {
          console.log("The column header was clicked.");
          //   spread.options.allowContextMenu = false;
        }
      });
      spread.bind(GC.Spread.Sheets.Events.EditStarting, function(sender, args) {
        var style = args.sheet.getStyle(args.row, args.col);
        console.log(style, "--style");
        if (
          style &&
          style.dropDowns &&
          style.dropDowns[0] &&
          (style.dropDowns[0].type ==
            GC.Spread.Sheets.DropDownType.dateTimePicker ||
            style.dropDowns[0].type ==
              GC.Spread.Sheets.DropDownType.monthPicker)
        ) {
          args.cancel = true;
        }
        if (args.sheetName === "HH" && args.col <= 17) {
          args.cancel = true;
        }
        if (args.sheetName === "PC" && args.col <= 20) {
          args.cancel = true;
        }
      });
      sheet.bind(GC.Spread.Sheets.Events.ClipboardPasting, function(
        sender,
        args
      ) {
        console.log("ClipboardPasting", args);
        if (args.cellRange.col <= 17) {
          args.cancel = true;
        }
      });
    },
    registCommand(spread) {
      let self = this;
      // 获取右键菜单数组
      var menuData = spread.contextMenu.menuData;

      console.log(menuData);
      //向上插入一行
      spread.commandManager().register("insertRowsBefore", {
        canUndo: true,
        execute: function(context, options, isUndo) {
          let Commands = GC.Spread.Sheets.Commands;
          // 在此加cmd
          options.cmd = "insertRowsBefore";
          if (isUndo) {
            Commands.undoTransaction(context, options);
            return true;
          } else {
            Commands.startTransaction(context, options);
            let sheet = spread.getActiveSheet();
            let sels = sheet.getSelections();
            console.log(sels, "--sels");
            if (sels && sels.length > 0) {
              for (let i = 0; i < sels.length; i++) {
                let sel = sels[i];
                let row = sel.row;
                let col = sel.col;
                let rowCount = sel.rowCount;
                let colCount = sel.colCount;
                sheet.addRows(row, rowCount);
                //下拉框
                // sheet.setCellType(
                //   row,
                //   2,
                //   self.customerGroupCombo,
                //   GC.Spread.Sheets.SheetArea.viewport
                // );
                //日期选择器
                // sheet.setStyle(row, 4, self.activityMonth);
                // sheet.setStyle(row, 5, self.posterDate);
                // sheet.setStyle(row, 6, self.posterDate);
                //行高
                self.vAlignCenter();
                self.hAlignCenter();
              }
            }
            Commands.endTransaction(context, options);
            return true;
          }
        }
      });
      //向下插入一行
      spread.commandManager().register("insertRowsBehind", {
        canUndo: true,
        execute: function(context, options, isUndo) {
          let Commands = GC.Spread.Sheets.Commands;
          // 在此加cmd
          options.cmd = "insertRowsBehind";
          if (isUndo) {
            Commands.undoTransaction(context, options);
            return true;
          } else {
            Commands.startTransaction(context, options);
            let sheet = spread.getActiveSheet();
            let sels = sheet.getSelections();
            console.log(sels, "--sels");
            // if (sels && sels.length > 0) {
            //   for (let i = 0; i < sels.length; i++) {
            //     let sel = sels[i];
            //     let row = sel.row;
            //     let col = sel.col;
            //     let rowCount = sel.rowCount;
            //     let colCount = sel.colCount;
            //     sheet.addRows(row + 1, rowCount);
            //     //下拉框
            //     sheet.setCellType(
            //       row + 1,
            //       2,
            //       self.customerGroupCombo,
            //       GC.Spread.Sheets.SheetArea.viewport
            //     );
            //     //日期选择器
            //     sheet.setStyle(row + 1, 4, self.activityMonth);
            //     sheet.setStyle(row + 1, 5, self.posterDate);
            //     sheet.setStyle(row + 1, 6, self.posterDate);
            //     //行高
            //     self.vAlignCenter();
            //     self.hAlignCenter();
            //   }
            // }
            Commands.endTransaction(context, options);
            return true;
          }
        }
      });
      // 定义一个在行头区域执行的右键菜单项
      var insertRowsBefore = {
        command: "insertRowsBefore",
        text: "向上插入行",
        // name只要不重复即可
        name: "insertRowsBefore",
        // 把自己定义好的icon class加在这里
        iconClass: "gc-spread-custom-icon",
        workArea: "rowHeader"
      };
      // 定义一个在行头区域执行的右键菜单项
      var insertRowsBehind = {
        command: "insertRowsBehind",
        text: "向下插入行",
        // name只要不重复即可
        name: "insertRowsBehind",
        // 把自己定义好的icon class加在这里
        iconClass: "gc-spread-custom-icon",
        workArea: "rowHeader"
      };

      // 将自定义的项，加入到“插入行”（insertRows）之后
      menuData.forEach(function(item, index) {
        if (item && item.name === "gc.spread.clearContents") {
          menuData.splice(index + 1, 0, { type: "separator" });
          menuData.splice(index + 2, 0, insertRowsBefore);
          menuData.splice(index + 3, 0, insertRowsBehind);
        }
      });
    },

    handleRemove() {},
    handleExceed() {},
    importExcel(event) {
      const file = event.target.files[0];
      let self = this;
      let excelIO = new ExcelIO.IO();
      excelIO.open(file, spreadJSON => {
        if (self.spread) {
          self.spread.fromJSON(spreadJSON);
        }
      });

      this.lockSpread(this.spread);
    },
    importSSExcel(file, fileList) {},
    exportExcel() {
      var serializationOption = {
        ignoreFormula: false,
        ignoreStyle: false,
        rowHeadersAsFrozenColumns: true,
        columnHeadersAsFrozenRows: true
      };
      if (this.spread) {
        var jsonStr = JSON.stringify(this.spread.toJSON(serializationOption));
        console.log(jsonStr);
        let excelIO = new ExcelIO.IO();
        excelIO.save(jsonStr, blob => {
          FileSaver.saveAs(blob, "export.xlsx");
        });
      }
    },
    fromJSON() {
      var jsonOptions = {
        ignoreFormula: false,
        ignoreStyle: false,
        frozenColumnsAsRowHeaders: false,
        frozenRowsAsColumnHeaders: false,
        doNotRecalculateAfterLoad: true
      };
      let jsonStr = ssJSON;
      if (this.spread) {
        this.spread.fromJSON(jsonStr, jsonOptions);
      }
    },
    lockSpread(spread) {
      if (spread) {
        console.log("lock");
        spread.options.newTabVisible = false;
        spread.options.allowContextMenu = false;
        let sheet = spread.getActiveSheet();
        sheet.options.isProtected = true;
      }
    },
    unlockSpread(spread) {
      if (spread) {
        console.log("unlock");
        spread.options.newTabVisible = true;
        spread.options.allowContextMenu = true;
        let sheet = spread.getActiveSheet();
        sheet.options.isProtected = false;
      }
    },
    download() {},
    upload() {
      var serializationOption = {
        ignoreFormula: true,
        ignoreStyle: true,
        includeBindingSource: true
      };
      if (this.spread) {
        var jsonStr = JSON.stringify(this.spread.toJSON(serializationOption));
        console.log(jsonStr);
      }

      //发送ajax请求将json数据传给服务器
    },
    hAlignCenter() {
      if (this.spread) {
        let sheet = this.spread.getSheet(0);
        let sheet_PC = this.spread.getSheet(1);
        let range = sheet.getSelections()[0];
        let range_PC = sheet_PC.getSelections()[0];

        sheet.suspendPaint();
        sheet_PC.suspendPaint();
        sheet
          .getRange(
            range.row,
            range.col,
            sheet.getRowCount(),
            sheet.getColumnCount()
          )
          .hAlign(GC.Spread.Sheets.HorizontalAlign.center);
        sheet_PC
          .getRange(
            range_PC.row,
            range_PC.col,
            sheet_PC.getRowCount(),
            sheet_PC.getColumnCount()
          )
          .hAlign(GC.Spread.Sheets.HorizontalAlign.center);
        sheet.resumePaint();
        sheet_PC.resumePaint();
      }
    },
    vAlignCenter() {
      if (this.spread) {
        let sheet = this.spread.getSheet(0);
        let sheet_PC = this.spread.getSheet(1);
        let range = sheet.getSelections()[0];
        let range_PC = sheet_PC.getSelections()[0];
        // let range = sheet.getRowCount;
        // let range = sheet.getColCount;
        let rowCount = sheet.getRowCount();
        let colCount = sheet.getColumnCount();
        let rowCount_PC = sheet_PC.getRowCount();
        let colCount_PC = sheet_PC.getColumnCount();
        sheet.suspendPaint();
        sheet_PC.suspendPaint();
        // for (let row = 0; row < rowCount; row++) {
        //   sheet.setRowHeight(row, 36);
        //   for (let col = 0; col < colCount; col++) {
        //     sheet
        //       .getCell(row, col)
        //       .cellPadding("0 0 0 0")
        //       .vAlign(GC.Spread.Sheets.VerticalAlign.center);
        //   }
        // }
        sheet.resumePaint();
        sheet_PC.resumePaint();
      }
    },
    initColumnWidth(sheet, sheet_PC) {
      sheet.suspendPaint();
      sheet_PC.suspendPaint();
      sheet.setColumnWidth(0, 100);
      sheet.setColumnWidth(1, 100);
      sheet.setColumnWidth(2, 500);
      sheet.setColumnWidth(3, 100);
      sheet.setColumnWidth(4, 100);
      sheet.setColumnWidth(5, 120);
      sheet.setColumnWidth(6, 120);
      sheet.setColumnWidth(7, 100);
      sheet.setColumnWidth(8, 100);
      sheet.setColumnWidth(9, 100);
      sheet.setColumnWidth(10, 120);
      sheet.setColumnWidth(11, 120);
      sheet.setColumnWidth(12, 200);
      sheet.setColumnWidth(13, 120);
      sheet.setColumnWidth(14, 300);
      sheet.setColumnWidth(15, 120);
      sheet.setColumnWidth(16, 120);
      sheet.setColumnWidth(17, 300);
      sheet.setColumnWidth(18, 120);
      sheet.setColumnWidth(19, 120);
      sheet.setColumnWidth(20, 200);
      sheet.setColumnWidth(21, 200); //Store Pool
      sheet.setColumnWidth(22, 200);
      sheet.setColumnWidth(23, 200);
      sheet.setColumnWidth(24, 120);
      sheet.setColumnWidth(25, 120);
      sheet.setColumnWidth(26, 120);
      sheet.setColumnWidth(27, 120);
      sheet.setColumnWidth(28, 120);
      sheet.setColumnWidth(29, 120);
      sheet.setColumnWidth(30, 120);
      sheet.setColumnWidth(31, 120);
      sheet.setColumnWidth(32, 120);
      sheet.setColumnWidth(33, 120);
      sheet.setColumnWidth(34, 120);
      sheet.setColumnWidth(35, 120);
      sheet.setColumnWidth(36, 120);
      sheet.setColumnWidth(37, 120);
      sheet.setColumnWidth(38, 120);
      sheet.setColumnWidth(39, 120);
      sheet.setColumnWidth(40, 120);
      sheet_PC.setColumnWidth(0, 100);
      sheet_PC.setColumnWidth(1, 100);
      sheet_PC.setColumnWidth(2, 500);
      sheet_PC.setColumnWidth(3, 100);
      sheet_PC.setColumnWidth(4, 100);
      sheet_PC.setColumnWidth(5, 120);
      sheet_PC.setColumnWidth(6, 120);
      sheet_PC.setColumnWidth(7, 100);
      sheet_PC.setColumnWidth(8, 100);
      sheet_PC.setColumnWidth(9, 100);
      sheet_PC.setColumnWidth(10, 120);
      sheet_PC.setColumnWidth(11, 120);
      sheet_PC.setColumnWidth(12, 200);
      sheet_PC.setColumnWidth(13, 120);
      sheet_PC.setColumnWidth(14, 300);
      sheet_PC.setColumnWidth(15, 120);
      sheet_PC.setColumnWidth(16, 120);
      sheet_PC.setColumnWidth(17, 120);
      sheet_PC.setColumnWidth(18, 120);
      sheet_PC.setColumnWidth(19, 120);
      sheet_PC.setColumnWidth(20, 500);
      sheet_PC.setColumnWidth(21, 120); //Store Pool
      sheet_PC.setColumnWidth(22, 120);
      sheet_PC.setColumnWidth(23, 120);
      sheet_PC.setColumnWidth(24, 120);
      sheet_PC.setColumnWidth(25, 120);
      sheet_PC.setColumnWidth(26, 120);
      sheet_PC.setColumnWidth(27, 120);
      sheet_PC.setColumnWidth(28, 120);
      sheet_PC.setColumnWidth(29, 120);
      sheet_PC.setColumnWidth(30, 120);
      sheet_PC.setColumnWidth(31, 120);
      sheet_PC.setColumnWidth(32, 120);
      sheet_PC.setColumnWidth(33, 120);
      sheet_PC.setColumnWidth(34, 120);
      sheet_PC.setColumnWidth(35, 120);
      sheet_PC.setColumnWidth(36, 120);
      sheet_PC.setColumnWidth(37, 120);
      sheet_PC.setColumnWidth(38, 120);
      sheet_PC.setColumnWidth(39, 120);
      sheet_PC.setColumnWidth(40, 120);
      sheet.resumePaint();
      sheet_PC.resumePaint();
    },
    disableContextMenu(spread) {
      function ContextMenu() {}
      ContextMenu.prototype = new GC.Spread.Sheets.ContextMenu.ContextMenu(
        spread
      );
      ContextMenu.prototype.onOpenMenu = (
        menuData,
        itemsDataForShown,
        hitInfo,
        spread
      ) => {
        var col = hitInfo.worksheetHitInfo.col;
        var hitType = hitInfo.worksheetHitInfo.hitTestType;
        if (col <= 17) {
          itemsDataForShown.splice(0, 20);
          var insertIndex = -1;
          var deleteIndex = -1;
          //   for (let i = 0; i < itemsDataForShown.length; i++) {
          //     var item = itemsDataForShown[i];
          //     if (item.name === "gc.spread.insertRows") {
          //       insertIndex = i;
          //     } else if (item.name === "gc.spread.deleteRows") {
          //       deleteIndex = i;
          //     }
          //   }
          //   if (insertIndex > -1) {
          //     itemsDataForShown.splice(insertIndex, 1);
          //   }
          //   if (deleteIndex > -1 && insertIndex > -1) {
          //     itemsDataForShown.splice(deleteIndex - 1, 1);
          //     itemsDataForShown.splice(deleteIndex - 2, 1);
          //   }
        }
      };
      spread.contextMenu = new ContextMenu();
    },
    setMultiColHeader(spread) {
      let sheet = spread.getSheet(0);
      let sheet_PC = spread.getSheet(1);
      let rowCount = sheet.getRowCount();
      let colCount = sheet.getColumnCount();
      let rowCount_PC = sheet_PC.getRowCount();
      let colCount_PC = sheet_PC.getColumnCount();
      sheet.suspendPaint();
      sheet_PC.suspendPaint();
      sheet.setRowCount(2, GC.Spread.Sheets.SheetArea.colHeader);
      sheet.setColumnCount(1, GC.Spread.Sheets.SheetArea.rowHeader);
      sheet.addSpan(0, 0, 1, 7, GC.Spread.Sheets.SheetArea.colHeader);
      sheet.setValue(0, 0, "门店信息", GC.Spread.Sheets.SheetArea.colHeader);
      sheet.addSpan(0, 10, 1, 4, GC.Spread.Sheets.SheetArea.colHeader);
      sheet.setValue(0, 10, "WSP信息", GC.Spread.Sheets.SheetArea.colHeader);
      sheet.addSpan(0, 14, 1, 3, GC.Spread.Sheets.SheetArea.colHeader);
      sheet.setValue(0, 14, "TCP", GC.Spread.Sheets.SheetArea.colHeader);
      sheet.addSpan(0, 18, 1, 2, GC.Spread.Sheets.SheetArea.colHeader);
      sheet.setValue(
        0,
        18,
        "2020 3月奥妙金纺钻石 AG",
        GC.Spread.Sheets.SheetArea.colHeader
      );
      sheet.addSpan(0, 20, 1, 4, GC.Spread.Sheets.SheetArea.colHeader);
      sheet.setValue(0, 20, "组套1", GC.Spread.Sheets.SheetArea.colHeader);
      sheet.addSpan(0, 24, 1, 4, GC.Spread.Sheets.SheetArea.colHeader);
      sheet.setValue(0, 24, "组套2", GC.Spread.Sheets.SheetArea.colHeader);
      sheet.addSpan(0, 28, 1, 3, GC.Spread.Sheets.SheetArea.colHeader);
      sheet.setValue(0, 28, "POSM信息", GC.Spread.Sheets.SheetArea.colHeader);
      sheet.addSpan(0, 31, 1, 3, GC.Spread.Sheets.SheetArea.colHeader);
      sheet.setValue(0, 31, "WSP信息", GC.Spread.Sheets.SheetArea.colHeader);
      sheet.addSpan(0, 34, 1, 3, GC.Spread.Sheets.SheetArea.colHeader);
      sheet.setValue(0, 34, "AG大陆演", GC.Spread.Sheets.SheetArea.colHeader);
      sheet.getRange(0, 0, 11, 7).backColor("#CCFFCC");
      sheet.getRange(0, 10, 11, 4).backColor("#ffd591");
      sheet.getRange(0, 14, 11, 3).backColor("#ffc53d");
      sheet.getRange(0, 18, 11, 2).backColor("#bae637");
      sheet.getRange(0, 20, 11, 4).backColor("#efdbff"); //
      sheet.getRange(0, 24, 11, 4).backColor("#36cfc9");
      sheet.getRange(0, 28, 11, 3).backColor("#69c0ff");
      sheet.getRange(0, 31, 11, 3).backColor("#2f54eb");
      sheet.getRange(0, 34, 11, 3).backColor("#eb2f96");

      sheet
        .getRange(0, 0, rowCount, colCount)
        .setBorder(
          new GC.Spread.Sheets.LineBorder(
            "gray",
            GC.Spread.Sheets.LineStyle.thin
          ),
          { all: true }
        );

      sheet_PC.setRowCount(2, GC.Spread.Sheets.SheetArea.colHeader);
      sheet_PC.setColumnCount(1, GC.Spread.Sheets.SheetArea.rowHeader);
      sheet_PC.addSpan(0, 0, 1, 7, GC.Spread.Sheets.SheetArea.colHeader);
      sheet_PC.setValue(0, 0, "门店信息", GC.Spread.Sheets.SheetArea.colHeader);
      sheet_PC.addSpan(0, 7, 1, 4, GC.Spread.Sheets.SheetArea.colHeader);
      sheet_PC.setValue(
        0,
        7,
        "201902 SSD Ave (kRMB)",
        GC.Spread.Sheets.SheetArea.colHeader
      );
      sheet_PC.addSpan(0, 11, 1, 4, GC.Spread.Sheets.SheetArea.colHeader);
      sheet_PC.setValue(0, 11, "WSP信息", GC.Spread.Sheets.SheetArea.colHeader);
      sheet_PC.addSpan(0, 15, 1, 5, GC.Spread.Sheets.SheetArea.colHeader);
      sheet_PC.setValue(0, 15, "TCP", GC.Spread.Sheets.SheetArea.colHeader);
      sheet_PC.addSpan(0, 20, 1, 1, GC.Spread.Sheets.SheetArea.colHeader);
      sheet_PC.setValue(
        0,
        20,
        "POSM历史信息",
        GC.Spread.Sheets.SheetArea.colHeader
      );
      sheet_PC.addSpan(0, 21, 1, 2, GC.Spread.Sheets.SheetArea.colHeader);
      sheet_PC.setValue(
        0,
        21,
        "2020 BPC 三八女神节 AG-Test",
        GC.Spread.Sheets.SheetArea.colHeader
      );
      sheet_PC.addSpan(0, 23, 1, 4, GC.Spread.Sheets.SheetArea.colHeader);
      sheet_PC.setValue(0, 23, "组套1", GC.Spread.Sheets.SheetArea.colHeader);
      sheet_PC.addSpan(0, 27, 1, 4, GC.Spread.Sheets.SheetArea.colHeader);
      sheet_PC.setValue(0, 27, "组套2", GC.Spread.Sheets.SheetArea.colHeader);
      sheet_PC.addSpan(0, 31, 1, 3, GC.Spread.Sheets.SheetArea.colHeader);
      sheet_PC.setValue(
        0,
        31,
        "POSM信息",
        GC.Spread.Sheets.SheetArea.colHeader
      );
      sheet_PC.addSpan(0, 34, 1, 2, GC.Spread.Sheets.SheetArea.colHeader);
      sheet_PC.setValue(0, 34, "WSP信息", GC.Spread.Sheets.SheetArea.colHeader);
      sheet_PC.addSpan(0, 36, 1, 3, GC.Spread.Sheets.SheetArea.colHeader);
      sheet_PC.setValue(
        0,
        36,
        "AG大陆演",
        GC.Spread.Sheets.SheetArea.colHeader
      );
      sheet_PC.getRange(0, 0, 11, 7).backColor("#CCFFCC");
      sheet_PC.getRange(0, 7, 11, 4).backColor("#ffd591");
      sheet_PC.getRange(0, 11, 11, 4).backColor("#ffc53d");
      sheet_PC.getRange(0, 15, 11, 5).backColor("#bae637");
      sheet_PC.getRange(0, 20, 11, 1).backColor("#95de64"); //
      sheet_PC.getRange(0, 21, 11, 2).backColor("#36cfc9");
      sheet_PC.getRange(0, 23, 11, 4).backColor("#69c0ff");
      sheet_PC.getRange(0, 27, 11, 4).backColor("#2f54eb");
      sheet_PC.getRange(0, 31, 11, 3).backColor("#eb2f96");
      sheet_PC.getRange(0, 34, 11, 2).backColor("#bae637");
      sheet_PC.getRange(0, 36, 11, 3).backColor("#ffd591");
      sheet_PC
        .getRange(0, 0, rowCount_PC, colCount_PC)
        .setBorder(
          new GC.Spread.Sheets.LineBorder(
            "gray",
            GC.Spread.Sheets.LineStyle.thin
          ),
          { all: true }
        );
      sheet.resumePaint();
      sheet_PC.resumePaint();
    },
    removeMenuDataItems(spread) {
      var menuData = spread.contextMenu.menuData;
      var newMenuData = [];
      menuData.forEach(item => {
        if (
          (item && item.name === "gc.spread.insertColumns") ||
          (item && item.name === "gc.spread.deleteColumns") ||
          (item && item.name === "gc.spread.insertRows") ||
          (item && item.name === "gc.spread.insertSheet") ||
          (item && item.name === "gc.spread.deleteSheet") ||
          (item && item.name === "gc.spread.hideSheet") ||
          (item && item.name === "gc.spread.unhideSheet")
        ) {
          return;
        }
        newMenuData.push(item);
      });
      console.log(newMenuData);
      spread.contextMenu.menuData = newMenuData;
    },
    setComboBox(spread) {
      let sheet = spread.getSheet(0);
      let sheet_PC = spread.getSheet(1);
      let rowCount = sheet.getRowCount();
      let colCount = sheet.getColumnCount();
      let rowCount_PC = sheet_PC.getRowCount();
      let colCount_PC = sheet_PC.getColumnCount();
      sheet.suspendPaint();
      sheet_PC.suspendPaint();

        sheet.setCellType(
          -1,
          18,
          this.storePoolGroupCombo,
          GC.Spread.Sheets.SheetArea.viewport
        );

        // sheet_PC.setCellType(
        //   -1,
        //   21,
        //   this.storePoolGroupCombo,
        //   GC.Spread.Sheets.SheetArea.viewport
        // );
      sheet.resumePaint();
      sheet_PC.resumePaint();
    },
    setDateStyle(spread) {
      let sheet = spread.getSheet(0);
      let sheet_PC = spread.getSheet(1);
      let rowCount = sheet.getRowCount();
      let colCount = sheet.getColumnCount();
      let rowCount_PC = sheet_PC.getRowCount();
      let colCount_PC = sheet_PC.getColumnCount();
      sheet.suspendPaint();
      sheet_PC.suspendPaint();

      sheet.setStyle(-1, 20, this.startDate);
      sheet.setStyle(-1, 21, this.endDate);
      sheet.setStyle(-1, 24, this.startDate);
      sheet.setStyle(-1, 25, this.endDate);
      sheet.setFormatter(-1, 20, "yyyy/mm/dd");
      sheet.setFormatter(-1, 21, "yyyy/mm/dd");
      sheet.setFormatter(-1, 24, "yyyy/mm/dd");
      sheet.setFormatter(-1, 25, "yyyy/mm/dd");

    //   sheet_PC.setStyle(-1, 23, this.startDate);
    //   sheet_PC.setStyle(-1, 24, this.endDate);
    //   sheet_PC.setStyle(-1, 27, this.startDate);
    //   sheet_PC.setStyle(-1, 28, this.endDate);
    //   sheet_PC.setFormatter(-1, 23, "yyyy/mm/dd");
    //   sheet_PC.setFormatter(-1, 24, "yyyy/mm/dd");
    //   sheet_PC.setFormatter(-1, 27, "yyyy/mm/dd");
    //   sheet_PC.setFormatter(-1, 28, "yyyy/mm/dd");
      sheet.resumePaint();
      sheet_PC.resumePaint();
    },
    setListStyle(spread) {
      let sheet = spread.getSheet(0);
      let sheet_PC = spread.getSheet(1);
      let rowCount = sheet.getRowCount();
      let colCount = sheet.getColumnCount();
      let rowCount_PC = sheet_PC.getRowCount();
      let colCount_PC = sheet_PC.getColumnCount();
      sheet.suspendPaint();
      sheet_PC.suspendPaint();

      sheet.setStyle(-1, 31, this.schedule);
    //   sheet_PC.setStyle(-1, 34, this.schedule);
      sheet.resumePaint();
      sheet_PC.resumePaint();
    },
    bindDataSource() {
      let sheet = new GC.Spread.Sheets.Worksheet("NewSheet");
      this.spread.addSheet(2, sheet);
      let activesheet = this.spread.getSheet(2);
      var nameColInfo = { name: "name", displayName: "姓名", size: 70 };
      var ageColInfo = {
        name: "age",
        displayName: "年龄",
        size: 40,
        resizable: false
      };
      var birthdayColInfo = {
        name: "birthday",
        displayName: "出生日期",
        formatter: "d/M/yy",
        size: 120
      };
      var positionColInfo = {
        name: "position",
        displayName: "职位",
        size: 50,
        visible: true
      };
      var isSaleColInfo = {
        name: "isSale",
        displayName: "是否促销",
        size: 50,
        visible: true,
        cellType: new GC.Spread.Sheets.CellTypes.CheckBox()
      };
      var colInfos = [
        { name: "name", displayName: "Name", size: 70 },
        { name: "age", displayName: "Age", size: 40, resizable: false },
        {
          name: "birthday",
          displayName: "Birthday",
          formatter: "d/M/yy",
          size: 120
        },
        { name: "position", displayName: "Position", size: 50, visible: false },
        {
          name: "isSale",
          displayName: "是否促销",
          size: 50,
          visible: true,
          cellType: new GC.Spread.Sheets.CellTypes.CheckBox()
        }
      ];
      sheet.autoGenerateColumns = false;
      activesheet.setDataSource(dataSource);
      //   activesheet.bindColumn(0, nameColInfo);
      //   activesheet.bindColumn(1, birthdayColInfo);
      //   activesheet.bindColumn(2, ageColInfo);
      //   activesheet.bindColumn(3, positionColInfo);
      //   activesheet.bindColumn(4, isSaleColInfo);
      activesheet.bindColumns(colInfos);
    }
  }
};
</script>
<style scoped>
/* *,
*:before,
*:after {
  -webkit-box-sizing: border-box !important;

  box-sizing: border-box !important;
} */
.home,
.spreadWrapper {
  height: calc(100% - 20px);
}
.spreadHost {
  height: calc(100% - 0px);
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
<style>
.el-card,
.el-card__body {
  height: 100% !important;
}
</style>