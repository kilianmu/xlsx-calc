"use strict";

const int_2_col_str = require('./int_2_col_str.js');
const col_str_2_int = require('./col_str_2_int.js');
const exec_formula = require('./exec_formula.js');
const find_all_cells_with_formulas = require('./find_all_cells_with_formulas.js');
const Calculator = require('./Calculator.js');

var mymodule = function(workbook, options) {
    var formulas = find_all_cells_with_formulas(workbook, exec_formula);
    const uniqueErrorMessages = new Set();

    for (var i = formulas.length - 1; i >= 0; i--) {
      try {
         // console.log(formulas[i].name);
        exec_formula(formulas[i]);
      } catch (error) {
        if (!options || !options.continue_after_error) {
          throw error
        }
        if (options.log_error) {
            console.log(error);
            const parts = error.message.split(':');
            const functionError = parts[parts.length - 1].trim();
            let errorMessage = "";
            if (functionError.includes("Function")){
                 errorMessage = `Error: ${functionError}`; //Sheet: ${formulas[i].sheet_name},
            } else {
                 errorMessage = `Error: ${functionError} - Sheet: ${formulas[i].sheet_name} - Cell ${formulas[i].name}`;
            }

            // If the error message is not in the uniqueErrorMessages Set, add it
            if (!uniqueErrorMessages.has(errorMessage)) {
                uniqueErrorMessages.add(errorMessage);
            }
           //console.log('error executing formula', 'sheet', formulas[i].sheet_name, 'cell', formulas[i].name, error.message)
        }
      }
    }
    if (uniqueErrorMessages.size > 0) {
        console.log('Unique errors executing formulas:');
        for (const errorMessage of uniqueErrorMessages) {
            console.log(errorMessage);
        }
    }
};

mymodule.calculator = function calculator(workbook) {
    return new Calculator(workbook, exec_formula);
};

mymodule.set_fx = exec_formula.set_fx;
mymodule.exec_fx = exec_formula.exec_fx;
mymodule.col_str_2_int = col_str_2_int;
mymodule.int_2_col_str = int_2_col_str;
mymodule.import_functions = exec_formula.import_functions;
mymodule.import_raw_functions = exec_formula.import_raw_functions;
mymodule.xlsx_Fx = exec_formula.xlsx_Fx;
mymodule.localizeFunctions = exec_formula.localizeFunctions;

mymodule.XLSX_CALC = mymodule

module.exports = mymodule;
