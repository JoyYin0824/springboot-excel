# springboot-excel
springboot 利用easyexcel 进行excel的导入和导出


1. 利用springboot构建构成
2. 整合easyexcel进行文件的导入和导出
3. 启用poi的导入(oom问题)


记录下问题:
  导入的时候因为excel的两种格式的问题出现了解析的问题，
  通过对流的封装解决 
  InputStream is = new BufferedInputStream(excl.getInputStream());
