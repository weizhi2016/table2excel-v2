# table2excel
text and image save to excel


### Installation

```
npm install js-table2excel-v2
```

### Usage
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
    type: 'test',
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
### table2excel options

|             | type   | default |
| ----------- | ------ | ------- |
| column      | Array  | []      |
| data        | Array  | []      |
| excelName   | String | -       |
| captionName | String | -       |



### column options

|       | introduction    | type   | default |
| ----- | --------------- | ------ | ------- |
| title | name for column | String | -       |
| key   | key for column  | String | -       |
| type  | text\|image     | String | text    |
