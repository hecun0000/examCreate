# 随机生成试题   

## 安装  
1. 安装node.js最新版，网址链接：[官网](https://nodejs.org/zh-cn/)   
2. 在当前目录下，打开命令行执行， `npm i`  


## 运行   

方法一： 

1. 在当前目录下，打开命令行执行， `npm run dev`  



方法二： 

1. 双击执行 `start.bat`  



方法三： 

1.  在当前目录下，打开命令行执行， `npm run build` 生成 `随机生成试卷.exe`  
2.  双击执行 `随机生成试卷.exe`


## 配置文件

配置文件为 'config.json' 

1. ratio: 题库中 专业题数目 / 通用题数目。  
   例子：   
     为0.4 则 专业题占40%， 通用题占60%。   
     为0 则 全部为通用题。
     为1 则 全部为专业题。
2. singleChoice： 单选题个数   
3. multipleChoice: 不定项选择题个数
4. booleanQuestion： 判断题个数