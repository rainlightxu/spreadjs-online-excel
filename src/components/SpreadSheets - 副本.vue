

<template>
  <div class="home">
    <div class="toolbar">
      <el-upload
        class="upload-demo"
        style="width:200px;float:left;"
        action
        :auto-upload="false"
        :on-remove="handleRemove"
        multiple
        :limit="1"
        :on-exceed="handleExceed"
        :on-change="importExcel()"
        ref="upload"
      > 
        <el-button size="small" type="primary">导入</el-button>
      </el-upload>
      <input type="file" class="el-button" @change="importExcel($event)" style="height: 40px;" />

      <el-button type="primary" size="small" @click="exportExcel">导出</el-button>
    </div>

    <div class="home">
      <div class="spreadWrapper">
        <gc-spread-sheets class="spreadHost" @workbookInitialized="workbookInitialized($event)">
          <gc-worksheet :dataSource="dataTable"></gc-worksheet>
        </gc-spread-sheets>
      </div>
    </div>
  </div>
</template> 
<script>
/* eslint-disable */
import Vue from "vue";

import { Message } from "element-ui";

import "@grapecity/spread-sheets/styles/gc.spread.sheets.excel2016colorful.css";
import "@grapecity/spread-sheets-vue";
import GC from "@grapecity/spread-sheets";
import ExcelIO from "@grapecity/spread-excelio";
import FileSaver from "file-saver";

export default {
  data() {
    return {
      fileList: [],
      autoGenerateColumns: true,
      width: 300,
      visible: true,
      resizable: true,
      formatter: "$ #.00",
      spread: null
    };
  },
  computed: {
    dataTable() {
      let dataTable = [];
      for (let i = 0; i < 42; i++) {
        dataTable.push({ price: i + 0.56 });
      }
      return dataTable;
    }
  },
  methods: {
    handleRemove() {},
    handleExceed() {},
    importExcel(event) {
      console.log(event);
      const e = event;
      const file = e.target.files[0];
      let self = this;
      let excelIO = new ExcelIO.IO();
      excelIO.open(file, spreadJSON => {
        if (self.spread) {
          self.spread.fromJSON(spreadJSON);
        }
      });
    },
    exportExcel() {
      let self = this;
      if (self.spread) {
        let spreadJSON = JSON.stringify(self.spread.toJSON());
        let excelIO = new ExcelIO.IO();
        excelIO.save(spreadJSON, blob => {
          FileSaver.saveAs(blob, "export.xlsx");
        });
      }
    },
    workbookInitialized(spread) {
      this.spread = spread;
    }
  }
};
</script>
<style scoped>
.home,
.spreadWrapper {
  height: calc(100% - 20px);
}
.spreadHost {
  width: 100%;
  height: 100%;
}
.toolbar {
  padding-bottom: 15px;
}
</style>