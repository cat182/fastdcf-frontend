import {errors} from "../functions/formulaParser"
import { cellRangeToArray,addCommas } from "./util"
import { computeExpression } from "./formulaParser"

const spreadsheetFunctions = {}

spreadsheetFunctions.SUM = (Arguments, entries, newComputedCells, callHistory, alreadyComputed) => {
    let sum = 0

    for(let a=0; a<Arguments.length; a++) {
        const argument = Arguments[a]
        let computedArgument = cellRangeToArray(entries, argument, newComputedCells, callHistory, alreadyComputed) // If it is false the argument does not represent a cell range

        // Pass down errors
        if (typeof(computedArgument) === "number") {
            return computedArgument
        }

        if (computedArgument === null) {
            computedArgument = computeExpression(entries, argument, newComputedCells, callHistory, alreadyComputed)

            // Pass down errors
            if (typeof(computedArgument) === "number") {
                return computedArgument
            }
        }

        if (Array.isArray(computedArgument)) {
            for (let v=0; v<computedArgument.length; v++) {
                let value = computedArgument[v]
                if (value === '') value = 0
                if (isNaN(value)) {
                    return errors.VALUE
                }

                sum += parseFloat(value)
            }
        } else {
            if (computedArgument === '') computedArgument = 0
            if (isNaN(computedArgument)) {
                return errors.VALUE
            }

            sum += parseFloat(computedArgument)
        }
    }

    return sum.toString()
}

spreadsheetFunctions.MULTIPLY = (Arguments, entries, newComputedCells, callHistory, alreadyComputed) => {
    let product = 1
    let count = 0

    for(let a=0; a<Arguments.length; a++) {
        const argument = Arguments[a]
        let computedArgument = cellRangeToArray(entries, argument, newComputedCells, callHistory, alreadyComputed) // If it is false the argument does not represent a cell range

        // Pass down errors
        if (typeof(computedArgument) === "number") {
            return computedArgument
        }

        if (computedArgument === null) {
            computedArgument = computeExpression(entries, argument, newComputedCells, callHistory, alreadyComputed)

            // Pass down errors
            if (typeof(computedArgument) === "number") {
                return computedArgument
            }
        }
        if (Array.isArray(computedArgument)) {
            for (let v=0; v<computedArgument.length; v++) {
                const value = computedArgument[v]
                if (value != '') {
                    if (isNaN(value)) {
                        return errors.VALUE
                    }

                    product *= parseFloat(value)
                    count++;
                }
            }
        } else {
            if (computedArgument != '') {
                if (isNaN(computedArgument)) {
                    return errors.VALUE
                }
                
                product *= parseFloat(computedArgument)
                count++;
            }
        }
    }

    if (count === 0) {
        return 0
    }

    return product.toString()
}

spreadsheetFunctions.AVERAGE = (Arguments, entries, newComputedCells, callHistory, alreadyComputed) => {
    let avg = 0
    let count = 0

    for(let a=0; a<Arguments.length; a++) {
        const argument = Arguments[a]
        let computedArgument = cellRangeToArray(entries, argument, newComputedCells, callHistory, alreadyComputed) // If it is false the argument does not represent a cell range

        // Pass down errors
        if (typeof(computedArgument) === "number") {
            return computedArgument
        }

        if (computedArgument === null) {
            computedArgument = computeExpression(entries, argument, newComputedCells, callHistory, alreadyComputed)

            // Pass down errors
            if (typeof(computedArgument) === "number") {
                return computedArgument
            }
        }
        if (Array.isArray(computedArgument)) {
            for (let v=0; v<computedArgument.length; v++) {
                const value = computedArgument[v]
                if (value != '""') {
                    if (isNaN(value)) {
                        return errors.VALUE
                    }

                    avg += parseFloat(value)
                    count += 1
                }
            }
        } else {
            if (computedArgument != '""') {
                if (isNaN(computedArgument)) {
                    return errors.VALUE
                }

                avg += parseFloat(computedArgument)
                count += 1
            }
        }
    }

    avg /= count
    if (count === 0) {
        return errors.DIV
    } else if (isNaN(avg)) {
        return errors.VALUE
    }

    return avg.toString()
}

spreadsheetFunctions.IF = (Arguments, entries, newComputedCells, callHistory, alreadyComputed) => {
    if (Arguments.length === 0 || Arguments.length > 3) return errors.NA

    const condition = Arguments[0]
    const ifTrue = Arguments[1]
    const ifFalse = Arguments[2]
    const computedCondition = computeExpression(entries, condition, newComputedCells, callHistory, alreadyComputed)
    
    // Return error
    if (typeof(computedCondition) === "number") {
        return computedCondition
    }

    if (condition === "") {
        return errors.NA
    } else if (computedCondition === '""' || computedCondition === "FALSE") {
        if (ifFalse) {
            return computeExpression(entries, ifFalse, newComputedCells, callHistory, alreadyComputed)
        } else {
            return '"FALSE"'
        }
    } else {
        if (computedCondition === "TRUE" || !isNaN(computedCondition)) {
            if (ifTrue) {
                return computeExpression(entries, ifTrue, newComputedCells, callHistory, alreadyComputed)
            } else {
                return '"TRUE"'
            } 
        }

        let notAString = true;
    
        if (computedCondition.startsWith('"') && computedCondition.endsWith('"')) {
            notAString = false;

            for (let i = 1; i < computedCondition.length - 1; i++) {
                const currentChar = computedCondition[i];
                const nextChar = computedCondition[i + 1];

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

        if (!notAString) {
            if (ifTrue) {
                return computeExpression(entries, ifTrue, newComputedCells, callHistory, alreadyComputed)
            } else {
                return '"TRUE"'
            }
        }

        return errors.VALUE
    }
}

function removeOuterQuotes(inputString) {
    if (inputString.startsWith('"') && inputString.endsWith('"')) {
        // Remove outer quotes
        return inputString.slice(1, -1);
    } else if (inputString.startsWith("'") && inputString.endsWith("'")) {
        // Remove outer quotes
        return inputString.slice(1, -1);
    } else {
        // No outer quotes found
        return inputString;
    }
}

spreadsheetFunctions.CONCAT = (Arguments, entries, newComputedCells, callHistory, alreadyComputed) => {
    let str = ""

    for(let a=0; a<Arguments.length; a++) {
        const argument = Arguments[a]
        let computedArgument = cellRangeToArray(entries, argument, newComputedCells, callHistory, alreadyComputed) // If it is false the argument does not represent a cell range

        // Pass down errors
        if (typeof(computedArgument) === "number") {
            return computedArgument
        }

        if (computedArgument === null) {
            computedArgument = computeExpression(entries, argument, newComputedCells, callHistory, alreadyComputed)

            // Pass down errors
            if (typeof(computedArgument) === "number") {
                return computedArgument
            }
        }

        if (Array.isArray(computedArgument)) {
            computedArgument.forEach((value) => {
                value = removeOuterQuotes(value)

                if (!isNaN(value) && value != "") {
                    value = addCommas(Math.round(1000*value)/1000)
                }

                str += value
            })
        } else {
            computedArgument = removeOuterQuotes(computedArgument)

            if (!isNaN(computedArgument) && computedArgument != "") {
                computedArgument = addCommas(Math.round(1000*computedArgument)/1000)
            }

            str += computedArgument
        }
    }

    return '"'+ str+'"'
};

spreadsheetFunctions.COUNTA = (Arguments, entries, newComputedCells, callHistory, alreadyComputed) => {
    let count = 0

    for(let a=0; a<Arguments.length; a++) {
        const argument = Arguments[a]
        let computedArgument = cellRangeToArray(entries, argument, newComputedCells, callHistory, alreadyComputed) // If it is false the argument does not represent a cell range

        // Pass down errors
        if (typeof(computedArgument) === "number") {
            return computedArgument
        }

        if (computedArgument === null) {
            computedArgument = computeExpression(entries, argument, newComputedCells, callHistory, alreadyComputed)

            // Pass down errors
            if (typeof(computedArgument) === "number") {
                return computedArgument
            }
        }
        if (Array.isArray(computedArgument)) {
            for (let v=0; v<computedArgument.length; v++) {
                const value = computedArgument[v]
                if (value != '""') {
                    count += 1
                }
            }
        } else {
            if (computedArgument != '""') {
                count += 1
            }
        }
    }

    return count.toString()
}

export default spreadsheetFunctions