# xoffice-parent
该项目是用来预览各类office文档的，目前可以将office文档转换成pdf或者html格式进行预览，还在不断的更新中，可以通过其他方式进行预览

  一、请求URL：
      1、http://ip:port/xoffice/xoffice?_key=false&_xformat=doc&_format=html&_file=需要预览的文档链接

      参数说明：(1)_key 是否清除服务器上的缓存文件，因为在预览的过程中，会产生用来预览的pdf或html等文件，可以提高访问的效率，
                  所以会涉及到缓存文件的参数传递，默认不传为true表示删除缓存文件，false为不删除
               (2)_xformat 需要预览的源文件的文件格式，例如：doc、docx、xls、xlsx、ppt、pptx、pdf、pdfx等，该参数是必传的
               (3)_format 以什么形式来进行预览，可以为pdf、html等，如果是以pdf文档的方式来进行预览，该项目中的操作是返回生成的pdf文档的流，
                  采用流的方式进行预览，如果是html格式，则会进行重定向的操作，跳转到生成的html文件的目录
               (4)_file 需要预览的文档的url连接，采用http协议进行获取，例如：http://localhost/resources/xxxx.doc
     
      2、http://ip:port/xoffice/oneKeyAct
          
          该url是用来一键生成预览文件的接口，可以在用户在公司app中提交需要预览的xoffice文件后，还没进行预览操作时，可以通过该url进行一键生成
          用来预览的文件，以提高用户体验流畅性
   
   二、定时任务的使用
        该项目采用到了spring自身的定时任务，用来周期性的生成用来预览的文件
