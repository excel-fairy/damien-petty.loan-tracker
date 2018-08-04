/**
 * Convert excel and google sheet column number to column name with letters
 * @param column column number
 * @return {string} column name
 * ex: columnToLetter(1) === 'A'
 * ex: columnToLetter(4) === 'D'
 */
function columnToLetter(column) {
    var temp, letter = '';
    while (column > 0) {
        temp = (column - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        column = (column - temp - 1) / 26;
    }
    return letter;
}

/**
 * Convert excel and google sheet column name with letters to column number
 * @param letter column name
 * @return {int} column number
 * ex: letterToColumn('A') === 1
 * ex: letterToColumn('D') === 4
 */
function letterToColumn(letter) {
    var column = 0, length = letter.length;
    for (var i = 0; i < length; i++) {
        column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
    }
    return column;
}

function letterToColumnStart0(letter) {
    return letterToColumn(letter) - 1;
}

function columnToLetterStart0(column) {
    return columnToLetter(column + 1);
}
