/** @OnlyCurrentDoc */
var g_objSheet = SpreadsheetApp.getActive();

const INPUT_ROW_FIRST = 2;
const OUTPUT_ROW_FIRST = 4;

const INPUT_FIELD_FLOOR = 1;
const INPUT_FIELD_ROOM = 2;
const INPUT_FIELD_IMP_DIM = 3;
const INPUT_FIELD_MET_DIM = 4;
const INPUT_FIELD_FLOORING = 5;

var g_intCurrentTableRow = OUTPUT_ROW_FIRST;
var g_intCurrentInputFieldInSet = 1;

function Process() {
  var objData = g_objSheet.getDataRange().getValues();
  var strRowValue;
  var strRowValueTemp;
  var arrSquareFeet;
  var i = 0;
  
  g_objSheet.getRange("D4:J1000").clearContent();

  objData.forEach(function(arrRow){
    strRowValue = arrRow[1];

    if(strRowValue != ""){
      if(i >= INPUT_ROW_FIRST){
        switch(g_intCurrentInputFieldInSet){
          case INPUT_FIELD_FLOOR:
            g_objSheet.getRange("D" + g_intCurrentTableRow).setValue(strRowValue);
            break;
          case INPUT_FIELD_ROOM:
            g_objSheet.getRange("E" + g_intCurrentTableRow).setValue(strRowValue);
            break;
          case INPUT_FIELD_IMP_DIM:
            g_objSheet.getRange("F" + g_intCurrentTableRow).setValue(strRowValue);
            
            strRowValueTemp = strRowValue;

            strRowValueTemp.replace(" p","");
            strRowValueTemp.replace(" p rr","");
            strRowValueTemp.replace(" ft","");
            strRowValueTemp.replace(" ft irr","");
            strRowValueTemp.replace(",",".");

            arrSquareFeet = strRowValueTemp.split(" X ");

            strRowValueTemp = parseFloat(arrSquareFeet[0]) * parseFloat(arrSquareFeet[1]);

            g_objSheet.getRange("I" + g_intCurrentTableRow).setValue(strRowValueTemp);
            g_objSheet.getRange("J" + g_intCurrentTableRow).setValue("");
            break;
          case INPUT_FIELD_MET_DIM:
            g_objSheet.getRange("G" + g_intCurrentTableRow).setValue(strRowValue);
            break;
          case INPUT_FIELD_FLOORING:
            g_objSheet.getRange("H" + g_intCurrentTableRow).setValue(strRowValue);
            break;
          default:
            break;
        }
        
        if(g_intCurrentInputFieldInSet<5){
          g_intCurrentInputFieldInSet++;
        }else{
          g_intCurrentInputFieldInSet = 1;
          g_intCurrentTableRow++;
        }    
      }
    }
    i++;
  });

  g_objSheet.getRange("B3:B1000").clearContent();
};