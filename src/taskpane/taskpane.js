/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("btn").addEventListener("click",writeData);
    document.getElementById("check").addEventListener("click",check)
    document.getElementById("reset").addEventListener("click",reset)
  }
});

export async function writeData() {
  try {
    await Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getActiveWorksheet()
      const range1 = ws.getRange("B1:B6")
      const range2 = ws.getRange("A1:A6")
      const range3 = ws.getRange("D1")
      const min = 0
      const max = 30
      const minCeiled = Math.ceil(min);
      const maxFloored = Math.floor(max);
      const rand_hours = Math.floor(Math.random()*(maxFloored - minCeiled) + minCeiled)
      const rand_gpa = (Math.random()+3).toFixed(2)
      range1.values = [["GPA"],[(Math.random()+3).toFixed(2)],[(Math.random()+3).toFixed(2)],[(Math.random()+3).toFixed(2)],[(Math.random()+3).toFixed(2)],[(Math.random()+3).toFixed(2)]]
      range2.values = [["Hours"],[Math.floor(Math.random()*(maxFloored - minCeiled) + minCeiled)],[Math.floor(Math.random()*(maxFloored - minCeiled) + minCeiled)],[Math.floor(Math.random()*(maxFloored - minCeiled) + minCeiled)],[Math.floor(Math.random()*(maxFloored - minCeiled) + minCeiled)],[Math.floor(Math.random()*(maxFloored - minCeiled) + minCeiled)]]
      range3.values = [["Correlation"]]
      context.sync()
    });
  } catch (error) {
    console.error(error);
  }
}

async function check() {
  await Excel.run(async (context) => {
    let ws = context.workbook.worksheets.getActiveWorksheet();
    let answer = ws.getRange("S6");
    let fmla = `=IF(ISFORMULA(D2),IF(FORMULATEXT(D2)<>"=CORREL(A2:A6,B2:B6)",
"Incorrect. The formula in D2 has an error.","Correct. The formula typed is:  =CORREL(A2:A6,B2:B6)"),
"Incorrect. Please type a formula in cell D2.")`;
    answer.values = [[fmla]];
    let response = ws.getRange("D4");
    response.values = [["=S6"]];
    await context.sync();
  });
}

async function reset() {
  await Excel.run(async (context) => {
    let ws = context.workbook.worksheets.getActiveWorksheet();
    let range = ws.getRange("D2:S6");
    range.clear();
    await context.sync();
  });
}