## 简介
该项目是一个纯 python 练手项目。
为了解决项目开发中，数据库字段经常变动而忘记修改表结构文档，文档与实际开发不一致

## 使用方式
在如下方法中输入数据库信息
```python
def getConnection(db):
    return pymysql.connect(host='localhost', port=3306, user='root', password='123456', db=db)

```
然后在 `mian` 方法里面输入你想生成 excel 的数据库：
```python
    dbList = ['test']
```
最后修改文本保存路径：
```python
wb.save('/Users/admin/Desktop/person/' + db + '.xlsx')
```

