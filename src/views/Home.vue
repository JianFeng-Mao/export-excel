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
        header-row-class-name="tableRow"
      >
        <el-table-column prop="orderNo" label="订单号" width="180">
        </el-table-column>
        <el-table-column prop="orderDate" label="下单时间" width="180">
        </el-table-column>
        <el-table-column prop="product" label="产品" width="180">
        </el-table-column>
        <el-table-column prop="norms" label="规格" width="100">
        </el-table-column>
        <el-table-column prop="price" label="单价" width="80">
        </el-table-column>
        <el-table-column prop="num" label="数量" width="50"> </el-table-column>
        <el-table-column prop="totalPrice" label="总价" width="80">
        </el-table-column>
        <el-table-column prop="name" label="姓名" width="80"> </el-table-column>
        <el-table-column prop="phone" label="号码" width="120">
        </el-table-column>
        <el-table-column prop="address" label="地址" width="80">
        </el-table-column>
        <el-table-column prop="status" label="状态" width="100">
        </el-table-column>
        <el-table-column prop="userName" label="达人名称" width="120">
        </el-table-column>
        <el-table-column prop="userId" label="达人ID" width="180">
        </el-table-column>
      </el-table>
    </div>
  </div>
</template>

<script>
import XLSX from 'xlsx';
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
      let inpData = this.exportData.split('\n订单号');
      inpData = inpData.map((item, index) => {
        if (index > 0) {
          return (item = '订单号' + item);
        }
        return item;
      });
      inpData.forEach((d) => {
        let arr = d.split(/\n/);
        // key：value形式的值
        let keyArr = [];
        // 普通文字字符串
        let arr2 = [];
        // key：value形式转为对象
        let keyObj = {};
        // 普通文字字符串转为对象
        let obj = {};

        for (let i = 0; i < arr.length; i++) {
          // 区分key：value格式和普通文字字符串，分别存入对应的数组中进行后续处理
          if (arr[i] === '') continue;
          if (arr[i].includes('：')) {
            keyArr.push(arr[i]);
          } else {
            arr2.push(arr[i]);
          }
        }
        // 按 ：（注意中英文符合）分割数组 [key, value]
        let keyArr2 = keyArr.map((item) => {
          return item.split('：');
        });
        // 按固定文案向keyObj中添加对应属性和值
        keyArr2.forEach((item) => {
          if (item[0].includes('订单号')) {
            keyObj.orderNo = item[1];
          } else if (item[0].includes('下单时间')) {
            keyObj.orderDate = item[1];
          } else if (item[0].includes('带货达人')) {
            let tempArr = item[1].split(' ');
            keyObj.userName = tempArr[0];
            keyObj.userId = tempArr[1];
          }
        });
        if (arr2[4] !== '极速退') {
          arr2.splice(4, 0, '极速退');
        }
        // 处理普通文字字符串， 必须按顺序输入
        let tempArr2 = arr2[14] ? arr2[14].split('，') : [];
        // 区分规格和数量，找到显示数量的下标，进行截取
        let normsNumReg = /\x(\d)+/;
        let normsNumIndex = arr2[1] && arr2[1].search(normsNumReg);
        let norms = arr2[1] && arr2[1].substring(0, normsNumIndex);
        let num = arr2[1] && arr2[1].substring(normsNumIndex);

        obj.product = arr2[0];
        obj.norms = norms;
        obj.num = num;
        obj.price = arr2[7];
        obj.totalPrice = arr2[12];
        obj.name = tempArr2.length > 0 ? tempArr2[0] : '-';
        obj.phone = tempArr2.length > 1 ? tempArr2[1] : '-';
        obj.address = tempArr2.length > 2 ? tempArr2[2] : '-';
        obj.status = arr2[15];
        this.tableData.push({ ...keyObj, ...obj });
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
          订单号: obj.orderNo,
          下单时间: obj.orderDate,
          产品: obj.product,
          规格: obj.norms,
          单价: obj.price,
          数量: obj.num,
          总价: obj.totalPrice,
          姓名: obj.name,
          号码: obj.phone,
          地址: obj.address,
          状态: obj.status,
          达人名称: obj.userName,
          达人ID: obj.userId,
        };
      });
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
