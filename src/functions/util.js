import { computeCell } from "./formulaParser";

export function linearTrend(data) {
    // Calculate the mean of x and y values
    const meanX = data.reduce((sum, point) => sum + point.x, 0) / data.length;
    const meanY = data.reduce((sum, point) => sum + point.y, 0) / data.length;

    // Calculate the slope (m) and y-intercept (b) using the least-squares method
    const numerator = data.reduce((sum, point) => sum + (point.x - meanX) * (point.y - meanY), 0);
    const denominator = data.reduce((sum, point) => sum + Math.pow(point.x - meanX, 2), 0);

    const slope = numerator / denominator;
    const intercept = meanY - slope * meanX;

    // Return the slope and intercept as an object
    return {
        slope: slope,
        intercept: intercept
    };
}

export function addCommas(number) {
    // Convert the number to a string
    let numberStr = number.toString();

    // Separate the integer and decimal parts
    let parts = numberStr.split(".");
    let integerPart = parts[0];
    let decimalPart = parts.length > 1 ? "." + parts[1] : "";

    // Check for negative sign
    let sign = "";
    if (integerPart[0] === '-') {
        sign = "-";
        integerPart = integerPart.slice(1); // Remove the negative sign for processing
    }

    // Initialize variables
    let result = "";
    let count = 0;

    // Iterate over the characters in reverse order for the integer part
    for (let i = integerPart.length - 1; i >= 0; i--) {
        // Add a comma every three digits
        if (count === 3) {
            result = "," + result;
            count = 0;
        }
        result = integerPart[i] + result;
        count++;
    }

    // Combine the sign, formatted integer, and decimal parts
    return sign + result + decimalPart;
}

export function isDigit(char) {
    // Use a regular expression to check if the character is a digit
    return /^\d$/.test(char);
  }
export function detectRepeatingSequence(arr) {
    const sequenceMap = new Map();
  
    for (let i = 0; i < arr.length; i++) {
      for (let j = i + 1; j < arr.length; j++) {
        const sequence = arr.slice(i, j + 1).toString();
  
        if (sequenceMap.has(sequence)) {
          const startIndex = sequenceMap.get(sequence);
          const repeatedSequence = arr.slice(startIndex, j + 1);
          return repeatedSequence;
        } else {
          sequenceMap.set(sequence, i);
        }
      }
    }
  
    return null; // No repeating sequence found
}

export const ascendingOrder = (value1, value2) => {
    return value1 > value2 ? [value2, value1] : [value1, value2]
}

export function isPercentage(str) {
    // Regular expression for a percentage (positive or negative) with optional string and spaces
    const percentageRegex = /^\s*-?\d+(\.\d+)?\s*(\w*\s*)?%\s*$/;

    // Test if the string matches the percentage format
    return percentageRegex.test(str);
}

export function excelColumnToNumber(column) {
    if (!(/^[A-Z]+$/.test(column))) {
        return null
    }

    let result = 0;

    for (let i = 0; i < column.length; i++) {
        const char = column[i].toUpperCase(); // Convert to uppercase for case-insensitivity
        const charCode = char.charCodeAt(0) - 'A'.charCodeAt(0) + 1;
        result = result * 26 + charCode;
    }

    return result;
}

export function numberToExcelColumn(columnNumber) {
    let header = "";
    while (columnNumber > 0) {
        let remainder = (columnNumber - 1) % 26; // Adjusting for 1-based indexing
        header = String.fromCharCode(65 + remainder) + header;
        columnNumber = Math.floor((columnNumber - 1) / 26);
    }
    return header;
}

export function parseCellRange(cellRange) {
    const regex = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/;

    const match = cellRange.match(regex);

    if (!match) {
        // Invalid cell range format
        return null;
    }

    const [, columnStart, rowStart, columnEnd, rowEnd] = match;

    return {
        rowStart: parseInt(rowStart),
        rowEnd: parseInt(rowEnd),
        columnStart : excelColumnToNumber(columnStart),
        columnEnd: excelColumnToNumber(columnEnd),
    };
}

export function cellRangeToArray(entries, cellRange, newComputedCells, callHistory, alreadyComputed) {
    const parsedCellRange = parseCellRange(cellRange);

    if (parsedCellRange) {
        const { rowStart, rowEnd, columnStart, columnEnd } = parsedCellRange
        const values = []
        for(let row=rowStart; row<=rowEnd; row++) {
            for(let column=columnStart; column<=columnEnd; column++) {
                //if (callHistory[callHistory.length-1][0] === row && callHistory[callHistory.length-1][1] === column) {
                //    return ERRORS["#REF!"]
                //}
                
                // Pass down the errors
                const computedCell = computeCell(entries, row, column, newComputedCells, callHistory, alreadyComputed)
                if (typeof(computedCell) === "number") {
                    return computedCell
                }

                values.push('"'+computedCell+'"')
            }
        }
        return values
    } else {
        return null
    }
}

export function isFirefox() {
    return navigator.userAgent.toLowerCase().includes('firefox')
}