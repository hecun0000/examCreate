const xlsx = require('node-xlsx')
const officegen = require('officegen')
const fs = require('fs')
const _ = require('lodash')

const path = require("path")

const configPath = path.join("./config.json")
console.log(configPath)
let isConfigExist = fs.existsSync(configPath)

let jsonConfig = null
if (isConfigExist) {
  jsonConfig = JSON.parse(fs.readFileSync(configPath, "utf8"))
} else {
  console.log('请填写配置信息！！！！')
}
console.log('读取配置信息： ', jsonConfig)

const { ratio,  ...order} = jsonConfig
console.log(order, 'order')
const {singleChoice, multipleChoice, booleanQuestion} = order

const sheets = xlsx.parse('./知识测试题库.xlsx')
const poolData = []
sheets.forEach(function (sheet) {
  const {
    name,
    list
  } = alterData(sheet)
  const questByTypeData = questionByType(list)
  poolData.push(questByTypeData)
});

const examQuestions = getTest(poolData, ratio)

// 生成文件
createWord(examQuestions, order)

// 转化数据
function alterData(sheet) {
  const list = []
  for (let i = 2; i < sheet["data"].length; i++) {
    const row = sheet['data'][i];
    if (row && row.length > 0) {
      const arr = row.slice(3, 9)
      list.push({
        number: row[0],
        content: row[1],
        type: row[2],
        ansArr: arr,
        ans: row[9],
        sheetName: sheet["name"]
      });
    }
  }
  return {
    name: sheet["name"],
    list
  }
}
// 生成word
function createWord(list, order) {
  let docx = officegen('docx')
  const date = new Date()
  const year = date.getFullYear()
  const month = date.getMonth() + 1
  const day = date.getDate()
  const time = date.getTime()

  docx.on('finalize', function (written) {
    console.log('试题已经生成，请查看！')
  })

  docx.on('error', function (err) {
    console.log(err)
  })
  const wordTitle = '知识测试题'
  // 生成标题
  createTitle(docx, wordTitle)
  // console.log(list)
  // 生成题目
  createTestByOrder(order, list, docx)

  // 导出word
  let out = fs.createWriteStream(wordTitle + year + '-' + month + '-' + day + '_' + time + '.docx')
  out.on('error', function (err) {
    console.log(err)
  })
  docx.generate(out)
}

// 按顺序生成题目
function createTestByOrder(order, list, docx) {
  const orderList = {
    singleChoice: 'single',
    multipleChoice: 'multiple',
    booleanQuestion: 'bool',
  }
  const numberTxt = ['一', '二', '三']
  const orderArr = []
  Object.keys(order).forEach(key => {
    const val = order[key]
    if (val > 0) {
      orderArr.push()
    }
  })
  const orderArrFilter = orderArr.filter(i => i > 0)
  .forEach((item, index) => {
    const numText = numberTxt[index]
    const type = orderList[item]
    switch (type) {
      case 'single':
        // 生成单选题
        if (list.single) {
          creatSingleQuestion(list.single, docx, numText)
        }
        break;
      case 'multiple':
        // 生成不定项选择题
        if (list.multiple) {
          creatMultipleQuestion(list.multiple, docx, numText)
        }
        break;
      case 'bool':
        // 生成判断题
        if (list.bool) {
          creatBoolQuestion(list.bool, docx, numText)
        }
        break;
      default:
        break;
    }
  })
}

// 生成单选题
function creatSingleQuestion(list, docx, numText) {
  const contentP = docx.createP()
  contentP.addText(numText + `、	单选题（每题1分，共${singleChoice}分，请将答案填写进下方表格中）`, {
    font_size: 12,
    bold: true,
    font_face: '微软雅黑',
    align: "center"
  })
  for (let i = 0; i < list.length; i++) {
    createOtherQuestion(list[i], docx, i)
  }
}
// 生成不定项选择题
function creatMultipleQuestion(list, docx, numText) {
  const contentP = docx.createP()
  contentP.addText(numText + `、	多选题（每题2分，共${multipleChoice * 2}分，请将答案填写进下方表格中）`, {
    font_size: 12,
    bold: true,
    font_face: '微软雅黑',
    align: "center"
  })
  for (let i = 0; i < list.length; i++) {
    createOtherQuestion(list[i], docx, i)
  }
}
// 生成判断题
function creatBoolQuestion(list, docx, numText) {
  const contentP = docx.createP()
  contentP.addText(numText + `、	判断题（每题1分，共${booleanQuestion}分，请填写“√”或“×”至下方表格）`, {
    font_size: 12,
    bold: true,
    font_face: '微软雅黑',
    align: "center"
  })
  for (let i = 0; i < list.length; i++) {
    createJudgeQuestion(list[i], docx, i)
  }
}

// 题目分类
function questionByType(list) {
  const res = {}
  for (let i = 0; i < list.length; i++) {
    const typeName = list[i].type
    const typeMap = {
      '单选题': 'single',
      '不定项选择题': 'multiple',
      '判断题': 'bool',
    }
    const type = typeMap[typeName]
    if (res[type]) {
      res[type].push(list[i])
    } else {
      res[type] = [list[i]]
    }
  }

  return res
}


// 生成标题
function createTitle(docx, wordTitle) {
  let pObj = docx.createP()
  pObj.addText(wordTitle, {
    font_size: 16,
    font_face: '微软雅黑',
    bold: true,
    align: "center"
  })
  pObj.options.align = 'center'
}

// 判断题生成
function createJudgeQuestion(list, docx, index) {
  const {
    content,
    ans
  } = list
  const contentP = docx.createP()
  contentP.addText((index + 1) + '、' + content, {
    font_size: 10,
    font_face: '微软雅黑',
    align: "center"
  })
  const ansTxt = ans === 'A' ? '√' : '×'
  createAns(ansTxt, docx)
  createSource(list, docx)
}

// 生成正确答案
function createAns(ans, docx) {
  const contentP = docx.createP()
  contentP.addText('正确答案: ' + ans, {
    font_size: 10,
    font_face: '微软雅黑',
    align: "center"
  })
}


// 生成题目来源
function createSource(list, docx) {
  const { sheetName, number } = list
  const contentP = docx.createP()
  contentP.addText('题目来源: ' + sheetName + '[' + number + ']', {
    font_size: 10,
    font_face: '微软雅黑',
    align: "center"
  })
}


// 选择题生成
function createOtherQuestion(list, docx, index) {
  const { content, type, ansArr, ans } = list
  const contentP = docx.createP()
  contentP.addText((index + 1) + '、' + content, {
    font_size: 10,
    bold: type !== '判断题',
    font_face: '微软雅黑',
    align: "center"
  })
  const optionsArr = ['A', 'B', 'C', 'D', 'E', 'F']
  for (let i = 0; i < ansArr.length; i++) {
    if (ansArr[i]) {
      const str = ansArr[i].toString()
      const strP = docx.createP()
      strP.addText(optionsArr[i] + '. ' + str, {
        font_size: 10,
        font_face: '微软雅黑',
        align: "center"
      })
    }
  }

  createAns(ans, docx)
  createSource(list, docx)
}

//从一个给定的数组arr中,随机返回num个不重复项
function getArrayItems(arr, num) {
  //新建一个数组,将传入的数组复制过来,用于运算,而不要直接操作传入的数组;
  var temp_array = new Array();
  for (var index in arr) {
    temp_array.push(arr[index]);
  }
  //取出的数值项,保存在此数组
  var return_array = new Array();
  for (var i = 0; i < num; i++) {
    //判断如果数组还有可以取出的元素,以防下标越界
    if (temp_array.length > 0) {
      //在数组中产生一个随机索引
      var arrIndex = Math.floor(Math.random() * temp_array.length);
      //将此随机索引的对应的数组元素值复制出来
      return_array[i] = temp_array[arrIndex];
      //然后删掉此索引的数组元素,这时候temp_array变为新的数组
      temp_array.splice(arrIndex, 1);
    } else {
      //数组中数据项取完后,退出循环,比如数组本来只有10项,但要求取出20项.
      break;
    }
  }
  return return_array;
}

// 产生题目个数
function getQuestionNumbers(nums, ratio) {
  const first = Math.floor(nums * ratio)
  return [first, nums - first]
}

// 随机生成题库
function getTest(poolData, ratio) {
  const questionMap = {
    single: getQuestionNumbers(singleChoice, ratio),
    multiple: getQuestionNumbers(multipleChoice, ratio),
    bool: getQuestionNumbers(booleanQuestion, ratio),
  }

  console.log(questionMap)

  const examPool = {
    single: createQuestionByType(poolData, questionMap.single, 'single'),
    multiple: createQuestionByType(poolData, questionMap.multiple, 'multiple'),
    bool: createQuestionByType(poolData, questionMap.bool, 'bool')
  }
  return examPool
}

// 生成类型题
function createQuestionByType(poolData, questionNums, type) {
  const [firstData, secondData] = poolData
  const [singleFirst, secondFirst] = questionNums
  const firstPool = getArrayItems(firstData[type], singleFirst)
  const secondPool = getArrayItems(secondData[type], secondFirst)
  const questionPool = [...firstPool, ...secondPool]
  return questionPool.length > 0 ? _.shuffle(questionPool) : null
}