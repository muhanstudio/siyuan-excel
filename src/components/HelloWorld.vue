<template>
    <div class="hello">
      <div style="position: absolute;top:0;">
          <input style="font-size:16px;" type="file" @change="uploadExcel" />
          <button style="font-size:16px; margin-left:10px; background-color: #42b983; color: white;" @click="dExcel">导出文件</button>
      </div>
  
      <div
        id="luckysheet"
        style="margin:0px;padding:0px;position:absolute;width:100%;left: 0px;top: 30px;bottom:0px;"
      ></div>
    </div>
  </template>  

<script>
import LuckyExcel from 'luckyexcel'
import { exportSheetExcel } from './exportSheet.js';
export default {
  name: 'HelloWorld',
  props: {
    msg: String
  },
  mounted() {
    // In some cases, you need to use $nextTick
    // this.$nextTick(() => {
          $(function () {
            luckysheet.create({
              container: 'luckysheet', // 设定DOM容器的id
              title: '思源Excel编辑器(muhanstudio&&terwer)', // 设定表格名称
              lang: 'zh', // 设定表格语言
              plugins:['chart'],
              data:[
                    {
                        "name": "Cell", //工作表名称
                        "color": "dark", //工作表颜色
                        "index": 0, //工作表索引
                        "status": 1, //激活状态
                        "order": 0, //工作表的下标
                        "hide": 0,//是否隐藏
                        "row": 36, //行数
                        "column": 18, //列数
                        "defaultRowHeight": 19, //自定义行高
                        "defaultColWidth": 73, //自定义列宽
                        "celldata": [], //初始化使用的单元格数据
                        "config": {
                            "merge":{}, //合并单元格
                            "rowlen":{}, //表格行高
                            "columnlen":{}, //表格列宽
                            "rowhidden":{}, //隐藏行
                            "colhidden":{}, //隐藏列
                            "borderInfo":{}, //边框
                            "authority":{}, //工作表保护
                            
                        },
                        "scrollLeft": 0, //左右滚动条位置
                        "scrollTop": 315, //上下滚动条位置
                        "luckysheet_select_save": [], //选中的区域
                        "calcChain": [],//公式链
                        "isPivotTable":false,//是否数据透视表
                        "pivotTable":{},//数据透视表设置
                        "filter_select": {},//筛选范围
                        "filter": null,//筛选配置
                        "luckysheet_alternateformat_save": [], //交替颜色
                        "luckysheet_alternateformat_save_modelCustom": [], //自定义交替颜色	
                        "luckysheet_conditionformat_save": {},//条件格式
                        "frozen": {}, //冻结行列配置
                        "chart": [], //图表配置
                        "zoomRatio":1, // 缩放比例
                        "image":[], //图片
                        "showGridLines": 1, //是否显示网格线
                        "dataVerification":{} //数据验证配置
                    },
                    {
                        "name": "Sheet2",
                        "color": "",
                        "index": 1,
                        "status": 0,
                        "order": 1,
                        "celldata": [],
                        "config": {}
                    },
                    {
                        "name": "Sheet3",
                        "color": "",
                        "index": 2,
                        "status": 0,
                        "order": 2,
                        "celldata": [],
                        "config": {},
                    }
                ]
            });
          });
      
    // });
  },
  methods:{
    uploadExcel(evt){
        const files = evt.target.files;
        if(files==null || files.length==0){
            alert("No files wait for import");
            return;
        }

        let name = files[0].name;
        let suffixArr = name.split("."), suffix = suffixArr[suffixArr.length-1];
        if(suffix!="xlsx"){
            alert("Currently only supports the import of xlsx files");
            return;
        }
        LuckyExcel.transformExcelToLucky(files[0], function(exportJson, luckysheetfile){
            
            if(exportJson.sheets==null || exportJson.sheets.length==0){
                alert("Failed to read the content of the excel file, currently does not support xls files!");
                return;
            }
            window.luckysheet.destroy();
            
            window.luckysheet.create({
                container: 'luckysheet', //luckysheet is the container id
                showinfobar:false,
                data:exportJson.sheets,
                title:exportJson.info.name,
                userInfo:exportJson.info.name.creator
            });
        });
    },
    selectExcel(evt){
        const value = this.selected;
        const name = evt.target.options[evt.target.selectedIndex].innerText;
        
        if(value==""){
            return;
        }

        LuckyExcel.transformExcelToLuckyByUrl(value, name, (exportJson, luckysheetfile) => {
            
            if(exportJson.sheets==null || exportJson.sheets.length==0){
                alert("Failed to read the content of the excel file, currently does not support xls files!");
                return;
            }
            
            window.luckysheet.destroy();
            
            window.luckysheet.create({
                container: 'luckysheet', //luckysheet is the container id
                showinfobar:false,
                data:exportJson.sheets,
                title:exportJson.info.name,
                userInfo:exportJson.info.name.creator
            });
        });
    },
    async dExcel() {
    // Get all sheets
    const luckysheet = window.luckysheet.getLuckysheetfile();
    exportSheetExcel(luckysheet);
},
  }
}
</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->
<style scoped>
h3 {
  margin: 40px 0 0;
}
ul {
  list-style-type: none;
  padding: 0;
}
li {
  display: inline-block;
  margin: 0 10px;
}
a {
  color: #42b983;
}
</style>
