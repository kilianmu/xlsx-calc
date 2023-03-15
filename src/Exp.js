"use strict";

const RawValue = require('./RawValue.js');
const Range = require('./Range.js');
const str_2_val = require('./str_2_val.js');
const dynamicArrayFormulas = require('./dynamic_array_formulas.js');

const MS_PER_DAY = 24 * 60 * 60 * 1000;

var exp_id = 0;

module.exports = function Exp(formula) {
    var self = this;
    self.id = ++exp_id;
    self.args = [];
    self.name = 'Expression';
    self.update_cell_value = update_cell_value;
    self.formula = formula;
    
    function update_cell_value() {
        try {
            if (Array.isArray(self.args) 
                    && self.args.length === 1
                    && self.args[0] instanceof Range) {
                throw Error('#VALUE!');
            }
            formula.cell.v = self.calc();
            formula.cell.t = getCellType(formula.cell.v);

            if (Array.isArray(formula.cell.v && formula.cell.name && formula.cell.f && formula.cell.f.match(new RegExp(dynamicArrayFormulas.join('|'), 'i')))) {
                const array = formula.cell.v;
                if (!validateResultMatrix(array)) {
                    throw new Error('#N/A');
                }

                const existingCell = formula.cell.name;
                const existingCellLetter = existingCell.match(/[A-Z]+/)[0];
                const existingCellNumber = existingCell.match(/[0-9]+/)[0];

                for (let i = 0; i < array.length; i++) {
                    const newCellNumber = parseInt(existingCellNumber) + i;

                    for (let j = 0; j < array[i].length; j++) {
                        const newCellValue = array[i][j];
                        let newCellType = getCellType(newCellValue);

                        // original cell
                        if (i === 0 && j === 0) {
                            formula.cell.v = newCellValue;
                            if (newCellType) formula.cell.t = newCellType;
                        } 
                        // other cells
                        else {
                            const newLetterIndex = existingCellLetter.charCodeAt(0) - 65 + j;
                            const newCellLetter = getCellLetter(newLetterIndex);

                            const newCell = newCellLetter + newCellNumber;
                            formula.sheet[newCell] = {
                                v: newCellValue,
                                t: newCellType,
                            };
                        }
                    }
                }
            }
        }
        catch (e) {
            var errorValues = {
                '#NULL!': 0x00,
                '#DIV/0!': 0x07,
                '#VALUE!': 0x0F,
                '#REF!': 0x17,
                '#NAME?': 0x1D,
                '#NUM!': 0x24,
                '#N/A': 0x2A,
                '#GETTING_DATA': 0x2B
            };
            if (errorValues[e.message] !== undefined) {
                formula.cell.t = 'e';
                formula.cell.w = e.message;
                formula.cell.v = errorValues[e.message];
            }
            else {
                throw e;
            }
        }
        finally {
            formula.status = 'done';
        }
    }

    function getCellLetter(columnIndex) {
        let newCellLetter = '';
        while (newLetterIndex >= 0) {
            newCellLetter = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'[newLetterIndex % 26] + newCellLetter;
            newLetterIndex = Math.floor(newLetterIndex / 26) - 1;
        }
    }

    function getCellType(cellValue) {
        if (typeof(cellValue) === 'string') {
            return 's';
        }
        else if (typeof(cellValue) === 'number') {
            return 'n';
        }
    }

    function validateResultMatrix(result) {
        // array must be greater than 0 and be symmetrical
        if (Array.isArray(result)) {
            for (let i = 0; i < result.length; i++) {
                if (!(result[i] instanceof Array)) {
                    return false;
                }
                if (result[i].length !== result[0].length) {
                    return false;
                }
            }
        }

        return true;
    }

    function isEmpty(value) {
        return value === undefined || value === null || value === "";
    }
    
    function checkVariable(obj) {
        if (typeof obj.calc !== 'function') {
            throw new Error('Undefined ' + obj);
        }
    }

    function getCurrentCellIndex() {
        return +self.formula.name.replace(/[^0-9]/g, '');
    }
    
    function exec(op, args, fn) {
        for (var i = 0; i < args.length; i++) {
            if (args[i] === op) {
                try {
                    if (i===0 && op==='+') {
                        checkVariable(args[i + 1]);
                        let r = args[i + 1].calc();
                        args.splice(i, 2, new RawValue(r));
                    } else {
                        checkVariable(args[i - 1]);
                        checkVariable(args[i + 1]);

                        let a = args[i - 1].calc();
                        let b = args[i + 1].calc();
                        if (Array.isArray(a)) {
                            a = a[getCurrentCellIndex() - 1][0];
                        }
                        if (Array.isArray(b)) {
                            b = b[getCurrentCellIndex() - 1][0];
                        }

                        let r = fn(a, b);
                        args.splice(i - 1, 3, new RawValue(r));
                        i--;
                    }
                }
                catch (e) {
                    // console.log('[Exp.js] - ' + formula.name + ': evaluating ' + formula.cell.f + '\n' + e.message);
                    throw e;
                }
            }
        }
    }

    function exec_minus(args) {
        for (var i = args.length; i--;) {
            if (args[i] === '-') {
                checkVariable(args[i + 1]);
                var b = args[i + 1].calc();
                if (i > 0 && typeof args[i - 1] !== 'string') {
                    args.splice(i, 1, '+');
                    if (b instanceof Date) {
                        b = Date.parse(b);
                        checkVariable(args[i - 1]);
                        var a = args[i - 1].calc();
                        if (a instanceof Date) {
                            a = Date.parse(a) / MS_PER_DAY;
                            b = b / MS_PER_DAY;
                            args.splice(i - 1, 1, new RawValue(a));
                        }
                    }
                    args.splice(i + 1, 1, new RawValue(-b));
                }
                else {
                    if (typeof b === 'string') {
                        throw new Error('#VALUE!');
                    }
                    args.splice(i, 2, new RawValue(-b));
                }
            }
        }
    }

    self.calc = function() {
        let args = self.args.concat();
        exec('^', args, function(a, b) {
            return Math.pow(+a, +b);
        });
        exec_minus(args);
        exec('/', args, function(a, b) {
            if (b == 0) {
                throw Error('#DIV/0!');
            }
            return (+a) / (+b);
        });
        exec('*', args, function(a, b) {
            return (+a) * (+b);
        });
        exec('+', args, function(a, b) {
            if (a instanceof Date && typeof b === 'number') {
                b = b * MS_PER_DAY;
            }
            return (+a) + (+b);
        });
        exec('&', args, function(a, b) {
            return '' + a + b;
        });
        exec('<', args, function(a, b) {
            return a < b;
        });
        exec('>', args, function(a, b) {
            return a > b;
        });
        exec('>=', args, function(a, b) {
            return a >= b;
        });
        exec('<=', args, function(a, b) {
            return a <= b;
        });
        exec('<>', args, function(a, b) {
            if (a instanceof Date && b instanceof Date) {
                return a.getTime() !== b.getTime();
            }
            if (isEmpty(a) && isEmpty(b)) {
                return false;
            }
            return a !== b;
        });
        exec('=', args, function(a, b) {
            if (a instanceof Date && b instanceof Date) {
                return a.getTime() === b.getTime();
            }
            if (isEmpty(a) && isEmpty(b)) {
                return true;
            }
            if ((a == null && b === 0) || (a === 0 && b == null)) {
                return true;
            }
            if (typeof a === 'string' && typeof b === 'string' && a.toLowerCase() === b.toLowerCase()) {
                return true;
            }
            return a === b;
        });
        if (args.length == 1) {
            if (typeof(args[0].calc) !== 'function') {
                return args[0];
            }
            else {
                return args[0].calc();
            }
        }
    };

    var last_arg;
    self.push = function(buffer) {
        if (buffer) {
            var v = str_2_val(buffer, formula);
            if (((v === '=') && (last_arg == '>' || last_arg == '<')) || (last_arg == '<' && v === '>')) {
                self.args[self.args.length - 1] += v;
            }
            else {
                self.args.push(v);
            }
            last_arg = v;
            //console.log(self.id, '-->', v);
        }
    };
};