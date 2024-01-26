import { evaluate } from 'mathjs'
import spreadsheetFunctions from "../functions/spreadsheetFunctions"
import { excelColumnToNumber } from './util';

export const errors = {
    REF: 0,
    NA: 1,
    VALUE: 2,
    DIV : 3,
    QUOTE : 4,
    NAME : 5
};

const arrayVersions = [];

export function computeCell(entries, row, column, newComputedCells, callHistory, alreadyComputed) {
    if (alreadyComputed[row] && alreadyComputed[row][column] != undefined) {
        return alreadyComputed[row][column]
    }

    for(let i=0; i<callHistory.length; i++) {
        if (row === callHistory[i][0] && column === callHistory[i][1]) {
            return errors.REF
        }
    }

    callHistory.push([row, column])

    let computedCell
    let entry = getCellValue(entries, row, column)
    entry = entry.replace(/\n/g, '')
    if (entry.startsWith("=")) {
        computedCell = computeExpression(entries, entry.substring(1, entry.length), newComputedCells, callHistory, alreadyComputed)

        if (typeof(computedCell) === "number") {
            if (newComputedCells[row] === undefined) {newComputedCells[row] = [] }
            if (computedCell === errors.REF) {
                newComputedCells[row][column] = "#REF!"
            } else if (computedCell === errors.NA) {
                newComputedCells[row][column] = "#N/A"
            } else if (computedCell === errors.VALUE) {
                newComputedCells[row][column] = "#VALUE"
            } else if (computedCell === errors.DIV) {
                newComputedCells[row][column] = "#DIV/0!"
            } else if (computedCell === errors.QUOTE) {
                newComputedCells[row][column] = "#QUOTE"
            } else if (computedCell === errors.NAME) {
                newComputedCells[row][column] = "#NAME"
            }
            alreadyComputed[row] = alreadyComputed[row] || []
            alreadyComputed[row][column] = alreadyComputed[row][column] || computedCell

            return computedCell
        }

        if (computedCell.startsWith('"') && computedCell.endsWith('"')) {
            computedCell = computedCell.substring(1, computedCell.length - 1)
        }

        computedCell = computedCell.replace(/""/g, '"')
    } else {
        computedCell = entry
    }

    if (newComputedCells[row] === undefined) {newComputedCells[row] = [] }
    newComputedCells[row][column] = computedCell

    alreadyComputed[row] = alreadyComputed[row] || []
    alreadyComputed[row][column] = alreadyComputed[row][column] || computedCell
    return computedCell
}

function checkQuoteEscaping(expression) {
    let quotes = false;
    for (let i = 0; i < expression.length; i++) {
        const currentChar = expression[i];
        const nextChar = expression[i + 1];
    
        if (quotes) {
            if (currentChar === '"' && nextChar === '"') {
                i++;
            } else if (currentChar === '"') {
                quotes = false;
            }
        } else {
            if (currentChar === '"' && !quotes) {
                quotes = true;
            }
        }
    }
    
    return !quotes
}

function removeDollars(expression) {
    let pos = 0
    let quotes = false
    while (pos < expression.length) {
        const char = expression[pos];
        const nextChar = expression[pos+1]

        if (!quotes && char === '"') {
            quotes = true;
        } else if (quotes && char === '"') {
            if (nextChar && nextChar === '"') {
                pos += 2;
                continue;
            } else {
                quotes = false;
            }
        }

        if (!quotes && char === "$") {
            expression = expression.substring(0, pos) + expression.substring(pos+1)
            continue
        }

        pos += 1;
    }
    
    return expression
}

function parseFunctions(entries, expression, newComputedCells, callHistory, alreadyComputed) {
    let quotes = false;
    let skip = false;
    let pos = 0;

    while (pos < expression.length) {
        if (skip) {
            skip = false;
            continue;
        }

        const char = expression[pos];
        const nextChar = expression[pos + 1];

        if (!quotes && char === '"') {
            quotes = true;
        } else if (quotes && char === '"') {
            if (nextChar && nextChar === '"') {
                skip = true;
                pos += 1;
                continue;
            } else {
                quotes = false;
            }
        }

        if (!quotes && char.match(/[a-z]/i)) {
            let functionNameStart = pos;
            let functionNameEnd;

            for (let i = pos; i < expression.length; i++) {
                const char = expression[i];

                // Detect cell references
                if (!isNaN(char)) {
                    break
                }

                if (char === '(') {
                    functionNameEnd = i;
                    break;
                }
            }

            if (functionNameEnd) {
                let functionName = expression.substring(functionNameStart, functionNameEnd);
                let functionEnd;
                let parenthesis = 0;
                const brackets = [];

                for (let i = functionNameEnd + 1; i < expression.length; i++) {
                    const char = expression[i];

                    if (parenthesis === 0) {
                        if (char === ";" || char == ",") {
                            brackets.push(i);
                        } else if (char === ")") {
                            functionEnd = i;
                            break;
                        }
                    }

                    if (char === '(') {
                        parenthesis += 1;
                    } else if (char === ")") {
                        parenthesis -= 1;
                    }
                }

                const Arguments = [];
                
                if (functionNameEnd !== functionEnd - 1) {
                    for (let argument = 1; argument <= brackets.length + 1; argument++) {
                        let start;
                        let end;
    
                        if (argument === 1) {
                            start = functionNameEnd + 1;
                        } else {
                            start = brackets[argument - 2] + 1;
                        }
    
                        if (argument === brackets.length + 1) {
                            end = functionEnd;
                        } else {
                            end = brackets[argument - 1];
                        }
    
                        Arguments.push(expression.substring(start, end).trim());
                    }
                }

                functionName = functionName.toUpperCase()
                const trueMatches = (functionName.match(new RegExp("TRUE", 'g')) || []).length
                const falseMatches = (functionName.match(new RegExp("FALSE", 'g')) || []).length
                if (spreadsheetFunctions[functionName]) {
                    const functionResult = spreadsheetFunctions[functionName](Arguments, entries, newComputedCells, callHistory, alreadyComputed);

                    if (typeof(functionResult) === "number") {
                        return functionResult
                    }

                    expression = expression.substring(0, functionNameStart) + functionResult + expression.substring(functionEnd + 1);
                    pos += functionResult.toString().length;
                } else if ((trueMatches === 0 && falseMatches === 0) || (falseMatches + trueMatches > 1)) {
                    return errors.NAME
                } else {
                    if (trueMatches === 1) {
                        pos += 4
                    } else if (falseMatches === 1) {
                        pos += 5
                    } else {
                        pos += 1
                    }
                }
            } else {
                pos += 1;
            }
        } else {
            pos += 1;
        }
    }

    return expression
}

function parseReferences(entries, expression, newComputedCells, callHistory, alreadyComputed) {
    let consecutiveCellReference = false
    let quotes = false;
    let skip = false;
    let pos = 0;

    while (pos < expression.length) {
        if (skip) {
            skip = false;
            continue;
        }

        const char = expression[pos];
        const nextChar = expression[pos + 1];
        let rowReferenceStart;

        if (!quotes && char === '"') {
            quotes = true;
        } else if (quotes && char === '"') {
            if (nextChar && nextChar === '"') {
                skip = true;
                pos += 1;
                continue;
            } else {
                quotes = false;
            }
        }

        if (!quotes && char.match(/[a-z]/i) && char === char.toUpperCase()) {
            rowReferenceStart = pos;
            let rowReferenceEnd;
            let columnReferenceStart;

            for (let i = pos; i < expression.length; i++) {
                if(expression[i] != " ") {
                    if (!isNaN(expression[i])) {
                        rowReferenceEnd = i;
                        columnReferenceStart = i;
                        break;
                    }
                }
            }

            if (rowReferenceEnd != null) {
                let columnReferenceEnd;

                for (let i = columnReferenceStart; i < expression.length; i++) {
                    if (isNaN(expression[i])) {
                        columnReferenceEnd = i;
                        break;
                    } else if (i === expression.length - 1) {
                        columnReferenceEnd = i + 1;
                        break;
                    }
                }

                if (columnReferenceEnd) {
                    if (consecutiveCellReference) {
                        return errors.VALUE
                    } else {
                        const col = excelColumnToNumber(expression.substring(rowReferenceStart, rowReferenceEnd));
                        const row = parseInt(expression.substring(columnReferenceStart, columnReferenceEnd));

                        if (col && !isNaN(expression.substring(columnReferenceStart, columnReferenceEnd)) ){
                            let cellValue = computeCell(entries, row, col, newComputedCells, callHistory, alreadyComputed)

                            // Pass down errors
                            if (typeof(cellValue) === "number") {
                                return cellValue
                            }

                            cellValue = cellValue.replaceAll('"', '""');

                            if (isNaN(cellValue) || cellValue === "") {
                                cellValue = '"' + cellValue + '"'
                            }

                            expression = expression.substring(0, rowReferenceStart) + cellValue + " " + expression.substring(columnReferenceEnd);
                            pos += cellValue.toString().length;
                            consecutiveCellReference = true
                            continue;
                        }
                    }
                }
            }
            pos += 1;
        } else {
            pos += 1;
        }

        if (char !== ':') {
            consecutiveCellReference = false
        }
    }

    return expression
}

function simplifyParentheses(entries, expression, newComputedCells, callHistory) {
    let pos = 0
    let stringStart = null
    let stringEnd = null
    while (pos < expression.length) {
        const char = expression[pos];
        const secondChar = expression[pos + 1];

        if (char === '"') {
            if (stringStart === null) {
                stringStart = pos;
            } else {
                if (secondChar === '"') {
                    pos += 2;
                    continue;
                } else {
                    stringEnd = pos;
    
                    // Remove outer parentheses
                    let iterations = 1;
                    while (
                        expression[stringStart-1] === '(' &&
                        expression[stringEnd+1] === ')'
                    ) {
                        expression =
                            expression.substring(0, stringStart - 1) +
                            expression.substring(stringStart, stringEnd + 1) +
                            expression.substring(stringEnd + 1 + 1);
                        iterations += 1;
                        stringStart-=1
                        stringEnd-=1
                    }
    
                    stringStart = null;
                    stringEnd = null;
                    pos -= iterations - 1
                }
            }
        }
        pos += 1;
    }
    
    return expression
}

function isText(expression) {
    expression = expression.trim()
    let notAString = true;
    
    if (expression.startsWith('"') && expression.endsWith('"')) {
        notAString = false;

        for (let i = 1; i < expression.length - 1; i++) {
            const currentChar = expression[i];
            const nextChar = expression[i + 1];

            if (currentChar === '"' && nextChar === '"') {
                // Skip the next character since it is an escaped double quote
                i++;
            } else if (currentChar === '"') {
                // If a single double quote is encountered without being escaped, set notAString to true
                notAString = true;
                break;
            }
        }
    }

    return !notAString
}

function computeConditional(operands, operators) {
    let result = true
    
    for(let i=0; i<operators.length; i++) {
        const operator = operators[i]

        if (operator == "==") {
            result = result && (operands[i] == operands[i+1])
        } else if (operator == "!=") {
            result = result && (operands[i] != operands[i+1])
        } else if (operator == ">") {
            result = result && (operands[i] > operands[i+1])
        } else if (operator == "<") {
            result = result && (operands[i] < operands[i+1])
        } else if (operator == ">=") {
            result = result && (operands[i] >= operands[i+1])
        } else if (operator == "<=") {
            result = result && (operands[i] <= operands[i+1])
        }

        if (result === false) {
            return false
        }
    }

    return true
}

function parseTextConditionals(expression) {
    let pos = 0
    let skip = false
    let inConditionalExpression = false
    let quotes = false
    let stringBuffer = ""
    let startingPos = null
    let operands = []
    let operators = []

    while(pos<expression.length) {
        let SKIP = skip
        if (skip) {
            skip = false;
        }

        let char = expression[pos];
        let nextChar = expression[pos + 1];

        if (!quotes && char === '"' && !SKIP) {
            quotes = true;
        } else if (quotes && char === '"') {
            if (nextChar && nextChar === '"') {
                skip = true;
            } else {
                quotes = false;
                if (inConditionalExpression) {
                    operands.push(stringBuffer.substring(1, stringBuffer.length))
                    expression = expression.substring(0, pos) + expression.substring(pos+1)
                    char=expression[pos]
                    nextChar=expression[pos+1]

                    stringBuffer = ""
                } else {
                    pos +=1;
                    continue
                }
            }
        }

        if (pos >= expression.length-1 && inConditionalExpression && !quotes) {
            const result = computeConditional(operands, operators)
            expression = expression.substring(0, startingPos) + result.toString() + expression.substring(startingPos)
            break
        }

        if (!quotes) {
            if ((char == "=" && nextChar == "=")
                || (char == "!" && nextChar == "=")
                || (char == "<")
                || (char == ">")
                || (char == "<" && nextChar == "=")
                || (char == ">" && nextChar == "=")
            ) {
                const addToPos = nextChar == "=" ? 2 : 1

                // Determine if there is a string to the left of operator
                let stringEnd = null
                for (let i=pos-1; i>=0; i--) {
                    if (expression[i] === '"') {
                        stringEnd = i
                        break
                    } else if (expression[i] !== ' ') {
                        break
                    }
                }
                for (let i=pos+addToPos; i<=expression.length; i++) {
                    if (expression[i] === '"') {
                        break
                    } else if (expression[i] !== " ") {
                        stringEnd = null
                        break
                    }
                }

                if (inConditionalExpression) {
                    if (char == "=" && nextChar == "=") {
                        expression = expression.substring(0, pos) + expression.substring(pos+2)
                        operators.push("==")
                    } else if (char == "!" && nextChar == "=") {
                        expression = expression.substring(0, pos) + expression.substring(pos+2)
                        operators.push("!=")
                    } else if (char == "<" && nextChar == "=") {
                        expression = expression.substring(0, pos) + expression.substring(pos+2)
                        operators.push("<=")
                    } else if (char == ">" && nextChar == "=") {
                        expression = expression.substring(0, pos) + expression.substring(pos+2)
                        operators.push(">=")
                    } else if (char == "<") {
                        expression = expression.substring(0, pos) + expression.substring(pos+1)
                        operators.push("<")
                    } else if (char == ">") {
                        expression = expression.substring(0, pos) + expression.substring(pos+1)
                        operators.push(">")
                    } else if (char != " ") {
                        inConditionalExpression = false
                    }

                    continue
                }
                
                // Determine if there is a string to the right of operator
                if (!inConditionalExpression && stringEnd != null) {
                    let skip2 = false
                    for(let i=stringEnd-1; i >= 0; i=i-1) {
                        if (skip2) {
                            skip2 = false;
                            continue;
                        }

                        if (expression[i] == '"') {
                            if (expression[i-1] == '"') {
                                skip2 = true;
                                continue
                            } else {
                                operands.push(expression.substring(i+1, stringEnd))
                                expression = expression.substring(0, i) + expression.substring(pos+addToPos)
                                startingPos = i
                                pos = i
                                break
                            }
                        }
                    }

                    inConditionalExpression = true
                    if (char == "=" && nextChar == "=") {
                        operators.push("==")
                    } else if (char == "!" && nextChar == "=") {
                        operators.push("!=")
                    } else if (char == "<" && nextChar == "=") {
                        operators.push("<=")
                    } else if (char == ">" && nextChar == "=") {
                        operators.push(">=")
                    } else if (char == "<") {
                        operators.push("<")
                    } else if (char == ">") {
                        operators.push(">")
                    } else if (char != " ") {
                        inConditionalExpression = false
                    }
                    continue
                }

                pos++;
            } else {
                if (char != " " && !quotes && startingPos != null && inConditionalExpression) {
                    const result = computeConditional(operands, operators)
                    expression = expression.substring(0, startingPos) + result.toString() + expression.substring(startingPos)
                    operands = []
                    operators = []
                    startingPos = null
                    
                    if (result) {
                        pos += 4
                    } else {
                        pos += 5
                    }

                    inConditionalExpression = false
                    continue
                }

                pos++; 
            }
        } else if (inConditionalExpression) {
            stringBuffer += char
            expression = expression.substring(0, pos) + expression.substring(pos+1)
        } else {
            pos++;
        }
    }

    return expression
}

export function computeExpression(entries, expression, newComputedCells, callHistory, alreadyComputed) {
    if (!checkQuoteEscaping(expression)) return errors.QUOTE

    expression = removeDollars(expression)
    expression = parseFunctions(entries, expression, newComputedCells, callHistory, alreadyComputed)
    //console.log("1: " + expression)
    if (typeof(expression) === "number") return expression
    expression = parseReferences(entries, expression, newComputedCells, callHistory, alreadyComputed)
    //console.log("2: " + expression)
    if (typeof(expression) === "number") return expression
    expression = simplifyParentheses(entries, expression, newComputedCells, callHistory)
    //console.log("3: " + expression)
    if (isText(expression)) return expression.trim()
    expression = parseTextConditionals(expression)
    //console.log("4: " + expression)

    if (expression === "") return ""

    //console.log(expression)
    try {
        const result = evaluate(expression
            .replaceAll(/""/g, '0')
            .replaceAll(/"false"/ig, "false")
            .replaceAll(/"true"/ig, "true")
            .replaceAll(/and/ig, "and")
            .replaceAll(/or/ig, "or")
        )

        if (result === true) {
            return "TRUE"
        } else if (result === false) {
            return "FALSE"
        }

        if (isNaN(result)) {
            return errors.DIV
        }

        return result.toString()
      } catch (error) {
        return errors.VALUE
      }
}

function getCellValue(entries, row, column) {
    if (entries[row] && entries[row][column]) {
        return entries[row][column]
    }

    return ""
}