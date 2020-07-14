<template>
  <div class="home">
    <el-container>
      <el-header>Header</el-header>
      <el-container>
        <el-aside width="200px">Aside</el-aside>
        <el-container>
          <el-main>
            <SpreadSheets />
          </el-main>
          <el-footer>Footer</el-footer>
        </el-container>
      </el-container>
    </el-container>
  </div>
</template>

<script>
// @ is an alias to /src
import SpreadSheets from "@/components/SpreadSheets";
export default {
  name: "Home",
  components: {
    SpreadSheets
  }
};
</script>
<style>
.el-header,
.el-footer {
  background-color: #b3c0d1;
  color: #333;
  text-align: center;
  line-height: 60px;
}
.el-aside {
  background-color: #d3dce6;
  color: #333;
  text-align: center;
  line-height: 200px;
}
.el-main {
  background-color: #e9eef3;
  color: #333;
}
body > .el-container {
  margin-bottom: 40px;
}
html,
body,
#app,
.home,
.el-container,
.el-main {
  height: 100%;
  margin: 0;
}
</style>