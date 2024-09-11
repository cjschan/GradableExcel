/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// eslint-disable-next-line @typescript-eslint/no-unused-vars
/* global console, document, Excel, Office */
const sample_size_min = 8; 
const sample_size_max =15;

const data_set =[["Sales","Advertising_Expense"],
[4881.4,107],
[6632.19,111],
[5468.94,122],
[5576.33,121],
[3561.08,72],
[5623.12,81],
[6404.84,110],
[5969.73,110],
[5601.71,110],
[9407.48,177],
[5764.94,111],
[7307.9,123],
[6114.7,119],
[5655.24,113],
[5224.37,94],
[5657.03,115],
[5546.4,85],
[5450.53,95],
[4398.32,90],
[6488.02,102],
[7006.22,146],
[4297.7,63],
[7221.91,114],
[4678.14,68],
[4221.71,91],
[7655.69,122],
[5920.92,101],
[4277.14,78],
[4162.6,86],
[6803.29,114],
[5134.79,85],
[5634.51,104],
[6586.94,101],
[5737.43,87],
[6367.04,143],
[5477.48,113],
[3243.33,59],
[6262.64,104],
[3744.31,87],
[6430.9,117],
[5512.28,84],
[6360.91,98],
[5684.65,110],
[6945.92,117],
[3616.88,76],
[6198.78,93],
[4206.41,91],
[4649.74,87],
[7689.81,135],
[4948.47,108],
[3807.75,75],
[6820.93,118],
[7744.53,142],
[6722.92,121],
[4817.69,70],
[5414.79,90],
[5864.35,125],
[4569.28,86],
[5036.13,109],
[7173.8,115],
[4973.73,81],
[5141.94,99],
[2634.11,35],
[5763.47,80],
[4403.44,95],
[4717.71,75],
[6734.07,133],
[4586.29,71],
[4045.33,91],
[6443.01,103],
[7879.81,129],
[4516.01,71],
[6915.27,123],
[6460.44,100],
[5195.44,80],
[4938.34,109],
[5746.28,104],
[5155.84,88],
[5388.71,101],
[5763.95,92],
[6657.98,102],
[7105.24,113],
[6668.87,132],
[3633.05,75],
[8256.47,143],
[4916.88,61],
[4543.24,97],
[6933.7,112],
[6790.71,106],
[5082.38,88],
[5465.86,96],
[4887.68,90],
[4022.41,88],
[5716.45,117],
[6578.99,107],
[3632.2,86],
[5857.49,118],
[5675.45,106],
[6402.19,116],
[6956.47,113],
[4511.34,83],
[4813.54,89],
[6643.42,115],
[5947.37,112],
[4933.16,100],
[5606.03,102],
[5587.16,126],
[4238.92,88],
[5849.64,111],
[5183.17,96],
[6379.39,96],
[6143.85,122],
[6280.8,117],
[6508.37,116],
[5330.23,126],
[5897.54,100],
[5233.72,114],
[6299.43,94],
[4956.22,106],
[5333.08,97],
[4492.06,102],
[7191.21,112],
[5109.75,84],
[7341.43,142],
[3602.24,80],
[5265.34,76],
[5871.97,123],
[7207.96,116],
[5413.7,112],
[5243.41,113],
[4308.95,100],
[4649.4,82],
[5807.66,102],
[4038.13,86],
[7023.35,120],
[5441.29,97],
[5055.21,83],
[5023.22,94],
[5445.08,108],
[5175.77,89],
[3994.41,84],
[4852.98,105],
[6585.88,105],
[6151.52,90],
[4756.39,91],
[5233.24,105],
[4460.93,71],
[4037.25,72],
[3846.87,86],
[4495.31,96],
[6197.66,106],
[7152.57,130],
[5146.48,117],
[4885.65,97],
[4926.22,100],
[5049.34,80],
[4243.53,100],
[4330.67,94],
[6624.4,106],
[4010.95,83],
[6215.92,110],
[6043.24,131],
[4564.85,98],
[5053.71,108],
[6667.25,114],
[6025.93,92],
[6529.36,104],
[5276.7,100],
[5935,102],
[4105.34,85],
[4895.31,100],
[6804.81,110],
[5436.79,129],
[5231.67,119],
[6395.49,143],
[3657.24,85],
[5408.6,117],
[5876.85,104],
[7089.87,144],
[5893.05,84],
[5577.26,83],
[4778.14,88],
[3256.24,58],
[4776.52,89],
[5133.34,85],
[5041.47,103],
[6068.33,107],
[6791.36,138],
[6335.4,119],
[4921.22,88],
[4262.07,82],
[7176.55,110],
[4994.3,74],
[6002.22,137],
[7476.9,124],
[5050.86,91],
[5131.61,66],
[7364.67,127],
[5331.58,98],
[5294.52,125,]];
var n = data_set.length-1;
// var last = n + 1
// each question is an array of arrays
const shade_cells = ["D2","D5:E5","D8"];
var questions = [[["Correlation between Sales and Advertising_Expense","D1","Calculate the correlation between Sales and Advertising_Expense in cell D2."],
["answer","D2",["=CORREL(A2:A_last,B2:B_last)","= CORREL(B2:B_last, A2:A_last)"]],
["Correct. ","Excel's CORREL() function calculates the correlation between two cell ranges. Since the correlation is form_val, a strong positive association exists between Sales and Advertising_Expense."],
[["CORREL(","(A2"],["Uses CORREL() but the cell range is wrong."],"FORMULA REQUIRED"],
["Incorrect. ",["The cell range for Sales is A2:A_last, and the cell range for Advertising_Expense is B2:B_last.","Use an Excel function to calculate the correlation."],"Use an Excel function to calculate the correlation."],
["Notes","(Order doesn't matter in CORREL)"]],
[["Linear regression model","D4","Fit a simple linear regression model to predict Sales based on Advertising_Expense in cell D5."],
["answer","D5",["=LINEST(A2:A_last,B2:B_last)"]],
["Correct. ","Excel's LINEST() function fits a simple linear regression model. The first cell range is the response variable, Sales, and the second cell range is the explanatory variable, Advertising_Expense."],
[["LINEST(","(A2"],["=LINEST(B2:B_last,A2:A_last)","Uses LINEST() but the cell range is wrong."],"FORMULA REQUIRED"],
["Incorrect. ",["The cell range for Sales is A2:A_last, and the cell range for Advertising_Expense is B2:B_last.","Excel's LINEST() function fits a simple linear regression model. The first cell range should be the response variable and the second cell range should be the explanatory variable."],"Use an Excel function to fit the simple linear regression model."],
["Notes","D5 should match the formula. D6 should match the value.","(Order matters in LINEST.)"]],
[["Predict Sales for Advertising_Expense = 110","D7","Predict Sales for a product with Advertising_Expense = 110 in cell D8."],
["answer","D8",["=D5*110+E5","=E5+D5*110"]],
["Correct. ","Cell D5 contains the regression model's slope, b_1, and cell E5 contains the intercept, b_0. So, the predicted Sales when Advertising_Expense = 110 is y_hat = b_0 + b_1 * (110)."],
[["110","D5"],["number rounded"],"FORMULA REQUIRED"], // check if 0 element here is in formula 1, then check format
["Incorrect. ",["Cell D5 contains the regression model's slope, b1, and cell E5 contains the intercept, b0.","Excel rounds the estimated regression coefficients. For better precision, use cell references instead of numerical values."],"Use an Excel formula with cell references to make the prediction."],
["Notes","(Correct formula, but uses numerical values instead of cell references.)","(Slope and intercepts are switched.)"]]];
// Set up questions and answers
var quest = "<ol>";
for (let i = 0; i < questions.length;i++){
  quest = quest + "<li>" +questions[i][0][2] + "</li>"
}
Office.onReady((info) => {
  quest += "</ol>";
  hide_answer_columns();
  if (info.host === Office.HostType.Excel) {  
    document.getElementById("btn").addEventListener("click", writeData);
    document.getElementById("check").addEventListener("click", check);
    document.getElementById("resetter").addEventListener("click", resetter);   
    document.getElementById("questions").innerHTML = quest; 
  }
});
export async function hide_answer_columns() {
  try {
    await Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getActiveWorksheet();
      let col_range = ws.getRange("R:T");
      col_range.columnHidden = true;
    });
  } catch (error) {
    console.error(error);
  }
}
export async function set_text() {
  try {
    await Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getActiveWorksheet();
      let range3;
      for (let i = 0; i < questions.length; i++){
        range3 = ws.getRange(questions[i][0][1]);
        range3.values = [[questions[i][0][0]]];
      }
    });
  } catch (error) {
    console.error(error);
  }
}
export async function writeData() {
  try {
    await Excel.run(async (context) => {
    resetter();
    set_text();
    // await context.sync();
    var samp = get_rand_data(true); // false gets all of data set
    await context.sync();
    const ws = context.workbook.worksheets.getActiveWorksheet();
    let last = samp.length + 1;
    let range1 = ws.getRange(`A1:A${last}`);
    let range2 = ws.getRange(`B1:B${last}`); 
    let temp1 = [];
    temp1.push([data_set[0][0]]);
    let temp2 = [];
    temp2.push([data_set[0][1]]);
    // Fill data columns A and B with random sample
    for (let i = 0; i < last - 1; i++){
      temp1.push([samp[i][0]]);
      temp2.push([samp[i][1]])
    };
    // document.getElementById("feedback").textContent = "len: " + last+ "  "+ n +data_set[0][0] +samp[0][0] +JSON.stringify(temp1);
    range1.values = temp1;
    range2.values = temp2;   
    range1.format.autofitColumns();
    range1.numberFormat = "0.00";
    range2.format.autofitColumns();
    for (let i = 0; i < shade_cells.length; i++){
      range1 = ws.getRange(shade_cells[i]);
      range1.format.fill.color = "#F9D8BC";
    }
  });
  } catch (error) {
    console.error(error);
  }
}
async function feedback_box_clear() {
  await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getActiveWorksheet().shapes;
    await context.sync();
    // Load all the shapes in the collection without loading their properties.
    shapes.load("items/$none");
    await context.sync();
    shapes.items.forEach(function (shape) {
        shape.delete();
    });  
    await context.sync();
  });
}
async function resetter() {
  await Excel.run(async (context) => {
    feedback_box_clear();
    let rng = "A1:G50" + data_set.length;
    let ws = context.workbook.worksheets.getActiveWorksheet();
    let range = ws.getRange(rng);
    range.clear();
    document.getElementById("feedback").textContent = "";
    // set_text();
    // writeData();
    await context.sync();
  });
}
function round_to(num,digits) {
  const mult = Math.pow(10,digits);
  return Math.round(num * mult)/mult;
}
function get_correlation_words(correl,response_var,explanatory_var) {
  let assoc = " positive";
  if (Math.abs(correl) < 0) {
    assoc = " negative";
  } // numbers from Investopedia
  if (Math.abs(correl) <= 0.3) {
    assoc = "weak" + assoc;
  } else if (Math.abs(correl) <= 0.75) {
    assoc = "moderate" + assoc;
  } else {
    assoc = "strong" + assoc;
  }
  return "Excel's CORREL() function calculates the correlation between two cell ranges." +
  `Since the correlation is ${correl}, a ${assoc} association exists between ${response_var} and ${explanatory_var}.`
}
async function check() {
  await Excel.run(async (context) => {
    feedback_box_clear()
    let regex_sp = /\s/gm;
    let regex = /_last/g;
    let ws = context.workbook.worksheets.getActiveWorksheet();   
    let text = [];
    const last = n + 2;
    await context.sync();
    let data = ws.getRange("A2");
    // let border_color = "#A00000" // red
    let shapes = context.workbook.worksheets.getActiveWorksheet().shapes;
    let textbox = [];
    let make_shape = true;
    // let number_answers = [];
    for (let i = 0; i < questions.length; i++){  
      if (make_shape) {
        let feedback_color = "#f2d7d5"; // pink
        // get feedback for textboxes
        let response = ws.getRange(questions[i][1][1]);
        // chk needed for exact calculations
        let chk = ws.getRange(questions[i][1][1].replaceAll("D","S"));
        chk.values =[[questions[i][1][2][0].replaceAll(regex,last).replaceAll("D","S").replaceAll("E5","T5")]];
        chk.load("values");
        response.load("formulas");
        response.load("values");
        data.load("values");
        await context.sync();
        let form_str = response.formulas[0].toString();
        let form_val = response.values[0];
        let ans_val = chk.values[0];
        await context.sync();
        form_str = form_str.replaceAll(regex_sp,'');
        text.push("");
        let correct_feedback = questions[i][2][1].replaceAll(/form_val/g,round_to(form_val,5));
        if (correct_feedback.indexOf('CORREL') >= 0) {
          correct_feedback = get_correlation_words(round_to(form_val,5),data_set[0][0],data_set[0][1])
        }
        if (i == 0 && data.values[0][0] == "") { // no data generated by user
            text[i] = `Select "Generate new data"`;
            feedback_color = "#e8e8e8"; // gray
            // border_color = "#000000"; // black
            make_shape = false;
        } else if (form_str == "") { // no answer entered by user
          text[i] = `${i+1}. ${questions[i][4][2]}`; // no input
          feedback_color = "#e8e8e8"; // gray
        } else if (JSON.stringify(form_str).includes("=")) {// is a formula
          if (JSON.stringify(form_str).includes(questions[i][3][0][0].replaceAll(regex,last))) {   
            if (form_val.toString() == ans_val.toString()) { // correct formula
              text[i] = `✅  ${i+1}. Correct. ` + correct_feedback;
              feedback_color = "#d4efdf"; // light green
            } else if (JSON.stringify(form_str).includes(questions[i][3][0][1].replaceAll(regex,last))) { // incorrect Excel cells     
            text[i] = `❌  ${i+1}. Incorrect. ${questions[i][4][1][0].replaceAll(regex,last)}`; // incorrect Excel cells 
            } else {
              text[i] = `❌  ${i+1}. Incorrect. ${questions[i][4][1][1]}`; // incorrect Excel function or rounding 
            }
          } else { // default
            text[i] = `❌  ${i+1}. Incorrect. ${questions[i][4][2]}`; // other wrong formula input
          }
        } else { // default
          text[i] = `❌  ${i+1}. Incorrect. ${questions[i][4][2]}`; // non-formula input
        }
        // text[i] += "form_val: "+form_val.toString() +"   ans_val: " + ans_val.toString() // test line
        // make textboxes
        textbox.push(shapes.addTextBox(text[i]));
        textbox[i].name = "Feedback"+(i+1);
        textbox[i].left = 500;
        textbox[i].top = 5 + 55*i; // 60
        textbox[i].height = 50;
        textbox[i].width = 350;
        textbox[i].fill.setSolidColor(feedback_color);
        // textbox[i].lineFormat.color = border_color;
      }
    }  
    await context.sync();
    let default_cell = ws.getRange("A1");
    default_cell.select();

    default_cell = ws.getRange("S1:T10"); // clear chk 
    default_cell.clear();
    default_cell = ws.getRange("A1");
    default_cell.select();
    // textbox.name = "Textbox";
    // let feedback = ws.getRange("D4");
    // feedback.values = text1;
    await context.sync();
  });
}
// This sample creates a rectangle positioned 100 pixels from the top and left sides
// of the worksheet and is 150x150 pixels.

/** data for instance of add-in */
// returns an array of indices for the sample
function random_sample(min,max,len){
  var arr = [];
  while(arr.length < len){
    var r = Math.floor(Math.random() * (max-min)) + min + 1;
      if(arr.indexOf(r) === -1) arr.push(r);
  }
  return arr;
}
// If true, returns a random sample from data using a random set of indices
function get_rand_data(get_samp){// false gets original array data
  var samp = [];
  if (get_samp) { // random sample of data_set
    n = random_sample(sample_size_min,sample_size_max,1)[0]
    var rand_samp = random_sample(0,data_set.length-1,n + 1)
    for (let i of rand_samp){
      samp.push(data_set[i]);
    } 
  } else { // all of data_set
    n = data_set.length - 1;
    return Array.from(data_set).slice(1); 
  }
  return samp;
}
