# excel-python-converter

A tiny tool for data conversion between excel and python



## Origin

During daily work, lots of data-process tasks given by `.xlsx\.xls` or other formats need to be imported  to your code, for example, python. Though there're much python modules which are completely qulified for such job, it takes a little time to write codes for importing. This tool help you importing data to your code faster and easiler.



## UI Setting
<img src="https://github.com/user-attachments/assets/3233692f-1181-4235-a742-0d5cfcb5401f" alt="GitHub Logo" width="700"/>

very simple, left part is for excel data and the right one for python code.  Introductions of each widget are as follows:

-  `Excel---->\<----Python`: convert data format between excel and python

- `with header`: decide if python code outputs with special varable name, the picture above choose `with header`

- `direction`: if you familiar with python, the relation of `vertical` and `horizontal` is similar to `data` and **list(zip(*data))**

<img src="https://github.com/user-attachments/assets/3233692f-1181-4235-a742-0d5cfcb5401f" alt="GitHub Logo" width="440"/>      vs     <img src="https://github.com/user-attachments/assets/9b4d5011-1562-47b6-8354-a7ec95188934" alt="GitHub Logo" width="450"/>

- `auto copy`： default yes, decide if output is directly copied to paste board 

- `number->str`: decide if a number such as 1 be output to "1" or 1

- `hierarchical`：for special excel data which has hierarchical stucture, output in **json** format
- <img src="https://github.com/user-attachments/assets/a3a599e8-3482-42e9-bc6c-eb4bd23caf0c" alt="GitHub Logo" width="250"/>    <img src="https://github.com/user-attachments/assets/a4525844-d6e2-4551-a214-6bd89e334ae3" alt="GitHub Logo" width="600"/>


## Other

- excel-python-converter can handle **line breaks in cells**
- 
  <img src="https://github.com/user-attachments/assets/9217e789-90e5-4c03-82dd-3b15529680c5" alt="GitHub Logo" width="200"/>    <img src="https://github.com/user-attachments/assets/12335149-abe0-40a7-9c28-4ed543fe3a00" alt="GitHub Logo" width="600"/>


