# 前端一键导出Excel表格
支持导出文本、单图、多图、合并 单元格等功能


### 安装

```
npm install js-table2excel-v2
```

### 使用示例
``` javascript
import table2excel from 'js-table2excel-v2'

const column = [
  {
    title: '产品名称',
    key: 'name',
  },
  {
    title: '主图',
    key: 'pic',
    type: 'image',//单图模式
  }, 
   {
    title: '图集',
    key: 'photos',
    type: 'images',// 多图模式
  }, 
   {
    title: '描述',
    key: 'remark',
    type: 'text',
    mergeOptions: { colspan: 2 }, // 合并2列
  },
   // 注意：被合并的列不需要再定义
]

const data = [
  {
    name: '智能手机',
    pic: 'photo1.jpg',
    photos: ["photo1.jpg", "photo2.jpg"],
    remark:'描述',
    size: [100, 60], // image size for picture
  },
  {
    name: '平板电脑',
    pic: 'photo1.jpg',
    photos: ["photo1.jpg", "photo2.jpg"],
    remark:'描述',
    mergeOptions: {
      name: { rowspan: 2 } // 合并2行
    }
  },
   {
    // name 被合并，不需要重复
    pic: 'photo1.jpg‘,
    photos: ["photo1.jpg", "photo2.jpg"],
    remark:'描述',
  }
]

table2excel({
  column: columns,
  data: data,
  excelName: '产品清单',
  captionName: '2023年产品目录'
})

```
### 参数说明

|             | type   | default |
| ----------- | ------ | ------- |
| column      | Array  | []      |
| data        | Array  | []      |
| excelName   | String | -       |
| captionName | String | -       |



### 表头参数

|       | introduction    | type   | default |
| ----- | --------------- | ------ | ------- |
| title | name for column | String | -       |
| key   | key for column  | String | -       |
| type  | text\image\images|   -  | String | text    |
| mergeOptions  | merge options |   -  | object | text    |
