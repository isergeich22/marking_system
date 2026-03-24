function findMatchesByPostingNumber(arr1, arr2) {
    const set2 = new Set(arr2.map(item => item.posting_number))
    return arr1.filter(item => set2.has(item.posting_number))
}

function splitArrayIntoChunks(arr, chunkSize = 100) {
    if (!Array.isArray(arr)) throw new Error("Входные данные должны быть массивом")
    if (arr.length <= chunkSize) return [arr]
    const result = []
    for (let i = 0; i < arr.length; i += chunkSize) {
        result.push(arr.slice(i, i + chunkSize))
    }
    return result
}

function compareStrings(str1, str2) {
    if (str1.length !== str2.length) {
        console.log('Строки разной длины!')
        return
    }
    for (let i = 0; i < str1.length; i++) {
        if (str1[i] !== str2[i]) {
            console.log(`❌ Различие на позиции ${i}: '${str1[i]}' (код ${str1.charCodeAt(i)}) vs '${str2[i]}' (код ${str2.charCodeAt(i)})`)
        }
    }
    console.log('✅ Если различий нет выше — строки идентичны по символам.')
}

module.exports = { findMatchesByPostingNumber, splitArrayIntoChunks, compareStrings }
