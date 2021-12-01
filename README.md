<h1 align="center">Welcome to generated test questions 👋</h1>
<p>
  <a href="https://www.npmjs.com/package/generated tes- questions" target="_blank">
    <img alt="Version" src="https://img.shields.io/npm/v/generated tes- questions.svg">
  </a>
  <a href="https://github.com/hecun0000/examCreate" target="_blank">
    <img alt="Documentation" src="https://img.shields.io/badge/documentation-yes-brightgreen.svg" />
  </a>
  <a href="https://github.com/hecun0000/examCreate/blob/master/LICENSE" target="_blank">
    <img alt="License: hecun" src="https://img.shields.io/badge/License-hecun-yellow.svg" />
  </a>
</p>

> 根据 excel 随机抽取一套试题，输出word文件，技术栈：node + node-xlsx + officegen + pkg。使用pkg将项目打包生成可执行程序。

### 🏠 [Homepage](https://github.com/hecun0000/examCreate)

## Install

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
## Author

👤 **hecun**

* Website: http://hecun.site
* Github: [@hecun0000](https://github.com/hecun0000)

## 🤝 Contributing

Contributions, issues and feature requests are welcome!<br />Feel free to check [issues page](https://github.com/hecun0000/examCreate/issues). 

## Show your support

Give a ⭐️ if this project helped you!

## 📝 License

Copyright © 2021 [hecun](https://github.com/hecun0000).<br />
This project is [hecun](https://github.com/hecun0000/examCreate/blob/master/LICENSE) licensed.
