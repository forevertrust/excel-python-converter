# excel-python-converter

A tiny tool for data conversion between excel and python



## Origin

During daily work, lots of data-process tasks given by `.xlsx\.xls` or other formats need to be imported  to your code, for example, python. Though there're much python modules which are completely qulified for such job, it takes a little time to write codes for importing. This tool help you importing data to your code faster and easiler.



## UI Setting

<img src="C:\Users\admin\AppData\Roaming\Typora\typora-user-images\image-20250227012718688.png" alt="image-20250227012718688" style="zoom:77%;" />

very simple, left part is for excel data and the right one for python code.  Introductions of each widget are as follows:

-  `Excel---->\<----Python`: convert data format between excel and python

- `with header`: decide if python code outputs with special varable name, the picture above choose `with header`

- `direction`: if you familiar with python, the relation of `vertical` and `horizontal` is similar to `data` and **list(zip(*data))**

  <img src="C:\Users\admin\AppData\Roaming\Typora\typora-user-images\image-20250227014320327.png" alt="image-20250227014320327" style="zoom:50%;" />   <img src="C:\Users\admin\AppData\Roaming\Typora\typora-user-images\image-20250227014337096.png" alt="image-20250227014337096" style="zoom:50%;" />

- `auto copy`： default yes, decide if output is directly copied to paste board 

- `number->str`: decide if a number such as 1 be output to "1" or 1

- `hierarchical`：for special excel data which has hierarchical stucture, output in **json** format

- ![image-20250227014736703](C:\Users\admin\AppData\Roaming\Typora\typora-user-images\image-20250227014736703.png)                  <img src="C:\Users\admin\AppData\Roaming\Typora\typora-user-images\image-20250227014717967.png" alt="image-20250227014717967" style="zoom:50%;" />

## Other

- excel-python-converter can handle **line breaks in cells**

![image-20250227015120218](C:\Users\admin\AppData\Roaming\Typora\typora-user-images\image-20250227015120218.png)                                       <img src="C:\Users\admin\AppData\Roaming\Typora\typora-user-images\image-20250227015053130.png" alt="image-20250227015053130" style="zoom:50%;" />

