<template>
  <div class="home">
    <div>
      <textarea
        name="excel"
        id="excel"
        cols="80"
        rows="20"
        v-model="exportData"
        placeholder="请输入数据"
      ></textarea>
    </div>
    <div class="btns">
      <el-button type="primary" @click="change">转换为表格数据</el-button>
      <el-button @click="exportBtn">导出excel</el-button>
      <el-button @click="clearData">清空数据</el-button>
      <el-button @click="clearTable">清空表格</el-button>
    </div>
    <div>
      <div>
        ----------------------共计{{
          tableData.length
        }}条数据-------------------------
      </div>
      <el-table
        :data="tableData"
        :border="true"
        :stripe="true"
        style="width: 100%"
        height="500"
        header-row-class-name="tableRow"
      >
        <el-table-column prop="product" label="产品" minWidth="180">
        </el-table-column>
        <el-table-column prop="name" label="姓名" minWidth="80">
        </el-table-column>
        <el-table-column prop="phone" label="号码" minWidth="120">
        </el-table-column>
        <el-table-column prop="address" label="地址" minWidth="80">
        </el-table-column>
      </el-table>
    </div>
  </div>
</template>

<script>
import * as XLSX from 'xlsx';
export default {
  name: 'Home',
  data() {
    return {
      exportData: '',
      tableData: [],
    };
  },
  methods: {
    change() {
      if (this.exportData === '') {
        alert('请输入要转换的数据');
        return;
      }
      let arr1 = this.exportData.split('订单编号');

      let arr2 = arr1.map((item) => {
        return item.split(/\n/);
      });

      let resArr = [];
      arr2 = arr2.map((item) => {
        return item.filter((i) => {
          return i !== '';
        });
      });
      arr2.forEach((arr, index) => {
        let obj = {};
        if (index > 0) {
          obj.product = arr[2];
          let i = arr.findIndex((item) => '在线支付'.includes(item));
          obj.name = arr[i + 3];
          obj.phone = arr[i + 4];
          obj.address = arr[i + 5];
          resArr.push(obj);
        }
        this.tableData = resArr;
      });
    },
    clearData() {
      this.exportData = '';
    },
    clearTable() {
      this.tableData = [];
    },
    workbook2blob(workbook) {
      // 生成excel的配置项
      var wopts = {
        // 要生成的文件类型
        bookType: 'xlsx',
        // // 是否生成Shared String Table，官方解释是，如果开启生成速度会下降，但在低版本IOS设备上有更好的兼容性
        bookSST: false,
        type: 'binary',
      };
      var wbout = XLSX.write(workbook, wopts);
      // 将字符串转ArrayBuffer
      function s2ab(s) {
        var buf = new ArrayBuffer(s.length);
        var view = new Uint8Array(buf);
        for (var i = 0; i !== s.length; ++i) view[i] = s.charCodeAt(i) & 0xff;
        return buf;
      }
      let buf = s2ab(wbout);
      var blob = new Blob([buf], {
        type: 'application/octet-stream',
      });
      return blob;
    },
    // 将blob对象 创建bloburl,然后用a标签实现弹出下载框
    openDownloadDialog(blob, fileName) {
      if (typeof blob === 'object' && blob instanceof Blob) {
        blob = URL.createObjectURL(blob); // 创建blob地址
      }
      var aLink = document.createElement('a');
      aLink.href = blob;
      // HTML5新增的属性，指定保存文件名，可以不要后缀，注意，有时候 file:///模式下不会生效
      aLink.download = fileName || '';
      var event;
      if (window.MouseEvent) event = new MouseEvent('click');
      //   移动端
      else {
        event = document.createEvent('MouseEvents');
        event.initMouseEvent(
          'click',
          true,
          false,
          window,
          0,
          0,
          0,
          0,
          0,
          false,
          false,
          false,
          false,
          0,
          null
        );
      }
      aLink.dispatchEvent(event);
    },
    exportBtn() {
      if (!this.tableData.length) {
        alert('表格数据为空！');
        return;
      }
      this.exportExcel();
    },
    exportExcel() {
      let sheet1Data = this.tableData.map((obj) => {
        return {
          产品: obj.product,
          姓名: obj.name,
          号码: obj.phone,
          地址: obj.address,
        };
      });
      console.log(XLSX);
      var sheet1 = XLSX.utils.json_to_sheet(sheet1Data);
      sheet1['!cols'] = [];
      for (const key in sheet1Data[0]) {
        if (Object.hasOwnProperty.call(sheet1Data[0], key)) {
          sheet1['!cols'].push({ wch: 550 });
        }
      }
      // 创建一个新的空的workbook
      var wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, sheet1, '订单');
      const workbookBlob = this.workbook2blob(wb);
      this.openDownloadDialog(workbookBlob, '订单.xlsx');
    },
  },
};
</script>

<style>
.home {
  margin: 40px;
}
.btns {
  margin: 20px 0;
}
.tableRow .el-table__cell {
  background-color: #7d7f81 !important;
  color: #fff;
}
</style>
