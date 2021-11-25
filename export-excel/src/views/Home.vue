<template>
  <div class="home">
    <div>
      <textarea
        name="excel"
        id="excel"
        cols="80"
        rows="20"
        v-model="exportData"
      ></textarea>
    </div>
    <div>
      <button @click="change">转换为表格数据</button>
    </div>
    <div>
      <el-table :data="tableData" style="width: 100%">
        <el-table-column prop="orderNo" label="订单号" width="100">
        </el-table-column>
        <el-table-column prop="orderDate" label="下单时间" width="100">
        </el-table-column>
        <el-table-column prop="product" label="产品" width="100">
        </el-table-column>
        <el-table-column prop="norms" label="规格" width="100">
        </el-table-column>
        <el-table-column prop="price" label="单价" width="80">
        </el-table-column>
        <el-table-column prop="num" label="数量" width="50"> </el-table-column>
        <el-table-column prop="totalPrice" label="总价" width="80">
        </el-table-column>
        <el-table-column prop="name" label="姓名" width="80"> </el-table-column>
        <el-table-column prop="phone" label="号码" width="100">
        </el-table-column>
        <el-table-column prop="price" label="实收金额" width="80">
        </el-table-column>
        <el-table-column prop="status" label="状态" width="100">
        </el-table-column>
        <el-table-column prop="userName" label="达人名称" width="100">
        </el-table-column>
        <el-table-column prop="userId" label="达人ID" width="100">
        </el-table-column>
      </el-table>
    </div>
  </div>
</template>

<script>
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
      let inpData = this.exportData.split('订单详情发货修改收货地址\n');
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
        // 处理普通文字字符串， 必须按顺序输入
        let tempArr2 = arr2[14] ? arr2[14].split('，') : [];

        // 区分规格和数量，找到显示数量的下标，进行截取
        let normsNumReg = /\x(\d)+/;
        let normsNumIndex = arr2[1].search(normsNumReg);
        let norms = arr2[1].substring(0, normsNumIndex);
        let num = arr2[1].substring(normsNumIndex);

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
  },
};
</script>
