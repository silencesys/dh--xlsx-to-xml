import xlsx from 'xlsx'

const file = xlsx.readFile('./FILENAME.xlsx')
const sheet = file.Sheets[file.SheetNames[0]]
const range = xlsx.utils.decode_range(sheet['!ref'])

const cleanStandardFormatting = (text) => {
  const defaultFormatting = text.replace(/(<span style="font-size:12pt;">)(.*?)(<\/span>)/g, '$2')
  return defaultFormatting
}

const appendTopicTags = (text) => {
  const defaultFormatting = text.replace(/(<span style="font-size:9pt;">)(.*?)(<\/span>)/g, '<topic>$2</topic>')
  return defaultFormatting
}

const appendAppendix = (text) => {
  const colonPosition = text.indexOf(': ')

  if (colonPosition === -1) {
    return text
  }

  const firstHalf = text.substring(0, colonPosition)
  const secondHalf = text.substring(colonPosition + 2)

  return `${firstHalf}<appendix>${secondHalf}</appendix>`
}

const appendAdditionalTags = (text) => {
  let cleanedText = cleanStandardFormatting(text)
  cleanedText = appendAppendix(cleanedText)
  cleanedText = appendTopicTags(cleanedText)

  return cleanedText
}

for (let rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
  const firstCell = sheet[xlsx.utils.encode_cell({ r: rowNum, c: 0 })]
  const secondCell = sheet[xlsx.utils.encode_cell({ r: rowNum, c: 1 })]

  console.log(`<entry>${appendAdditionalTags(firstCell.h)}</entry>`, `<entry>${appendAdditionalTags(secondCell.h)}</entry>`)
}
