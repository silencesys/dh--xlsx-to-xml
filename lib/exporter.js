import xlsx from 'xlsx'
import fs from 'fs'
import yargs from 'yargs'
import { hideBin } from 'yargs/helpers'
import { exit } from 'process'
import COLOR from '../utils/colors.js'

const argv = yargs(hideBin(process.argv))
  .option('input', {
    alias: 'i',
    describe: 'The input file to be converted to xlsx',
    type: 'string'
  })
  .option('output', {
    alias: 'o',
    describe: 'The output file to be created',
    type: 'string',
  })
  .option('config', {
    alias: 'c',
    describe: 'The config file to be used',
    type: 'string',
  })
  .help()
  .alias('help', 'h')
  .argv

if (!argv.input) {
  // Fail if no input file is set.
  console.info('Please provide an input file')
  exit(0)
}

// Set global variables
const INPUT_FILE = argv.input
const OUTPUT_FILE = argv.output || argv.input.replace('.xlsx', '.xml')
let CONFIG = null
if (argv.config) {
  CONFIG = JSON.parse(fs.readFileSync(argv.config))
}

/**
 * Remove pre-defined tags from the text.
 *
 * @param {string} text value to be cleaned
 * @returns
 */
const stripTags = (text) => {
  if (CONFIG.stripTags.length === 0) {
    return text
  }

  let cleanedText = text
  for (const tag of CONFIG.stripTags) {
    const startTag = tag
    const endTag = `</${tag.slice(1, tag.indexOf(' '))}>`
    const regex = new RegExp(`(${startTag})(.*?)(${endTag})`, 'g')

    cleanedText = cleanedText.replace(regex, '$2')
  }

  return cleanedText
}

/**
 * Replace pre-defined tags from the text.
 *
 * @param {string} text value to be cleaned
 * @returns
 */
const replaceTags = (text) => {
  if (CONFIG.replaceTags.length === 0) {
    return text
  }

  let cleanedText = text

  for (const tag of CONFIG.replaceTags) {
    const startTag = tag.from
    const endTag = `</${tag.from.slice(1, tag.from.indexOf(' '))}>`
    const replacementStartTag = tag.to
    const replacementEndTag = `</${tag.to.slice(1, tag.to.indexOf(' '))}>`
    const regex = new RegExp(`(${startTag})(.*?)(${endTag})`, 'g')
    cleanedText = cleanedText.replace(regex, `${replacementStartTag}$2${replacementEndTag}`)
  }

  return cleanedText
}

/**
 * Replace column value by specified dividers.
 *
 * @param {string} text value to be cleaned
 * @returns
 */
const divideBy = (text) => {
  if (CONFIG.divideBy.length === 0) {
    return text
  }

  for (const divider of CONFIG.divideBy) {
  const colonPosition = text.indexOf(divider[0])

  if (colonPosition === -1) {
    return text
  }

  const firstHalf = text.substring(0, colonPosition).trim()
  const secondHalf = text.substring(colonPosition + 2).trim()

  return `${firstHalf} <${divider[1]}>${secondHalf}</${divider[1]}>`
  }
}

const cleanAndFormatColumnValue = (text) => {
  try {
    let cleanedText = stripTags(text)
    cleanedText = replaceTags(cleanedText)
    cleanedText = divideBy(cleanedText)

    return cleanedText
  } catch (e) {
    console.error(`${COLOR.fgRed}${e.message}${COLOR.reset}\n`)
    console.info(`INFORMATION`)
    if (!CONFIG) {
      console.info('You might forgot to provide config file, see @TODO')
    } else {
      console.info(`Check the source file or config file for possible errors.`)
    }
    exit(1)
  }
}

// Read the input file
try {
  const file = xlsx.readFile(INPUT_FILE)
  const sheet = file.Sheets[file.SheetNames[0]]
  const range = xlsx.utils.decode_range(sheet['!ref'])
  const code = []

  for (let rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
    const firstCell = sheet[xlsx.utils.encode_cell({ r: rowNum, c: 0 })]
    const secondCell = sheet[xlsx.utils.encode_cell({ r: rowNum, c: 1 })]

    if (firstCell && secondCell) {
      const idFirstCell = `${firstCell.v.slice(0, 1).toUpperCase()}-0-${rowNum}`
      const idSecondCell = `${firstCell.v.slice(0, 1).toUpperCase()}-1-${rowNum}`

      let additionToSecondCell = ''
      if (sheet[xlsx.utils.encode_cell({ r: rowNum + 1, c: 1 })]) {
        const thirdCell = sheet[xlsx.utils.encode_cell({ r: rowNum + 1, c: 1 })]
        if (thirdCell.v.length > 0) {
          additionToSecondCell = cleanAndFormatColumnValue(thirdCell.h)
        }
      }
      const firstColumnLanguage = CONFIG.language.length > 0 ? `xml:lang="${CONFIG.language[0]}"` : ''
      const secondColumnLanguage = CONFIG.language.length > 1 ? `xml:lang="${CONFIG.language[1]}"` : ''
      code.push(`    <${CONFIG.rowTagName} id="${idFirstCell}" corresp="${idSecondCell}" ${firstColumnLanguage}>${cleanAndFormatColumnValue(firstCell.h)}</${CONFIG.rowTagName}>\n    <${CONFIG.rowTagName} id="${idSecondCell}" corresp="${idFirstCell}" ${secondColumnLanguage}>${cleanAndFormatColumnValue(secondCell.h) + additionToSecondCell}</${CONFIG.rowTagName}>\n`)
    } else if (secondCell) {
      console.log(`${COLOR.fgYellow}Row ${rowNum + 1} has only one cell.${COLOR.reset}`)
      console.log(`${COLOR.fgYellow}Cell content was merged with previous cell.${COLOR.reset}`)
    } else {
      console.warn(`${COLOR.fgRed}Row ${rowNum + 1} is empty${COLOR.reset}`)
    }
  }

  // Create the output file
  let xml = `<?xml version="1.0" encoding="UTF-8" ?>\n  <${CONFIG.parentTagName}>\n`
  code.map(line => {
    xml += line
  })
  xml += `  </${CONFIG.parentTagName}>\n</xml>`

  fs.writeFileSync(OUTPUT_FILE, xml)
  console.info(`${COLOR.bgGreen}${COLOR.fgBlack}Created ${OUTPUT_FILE}${COLOR.reset}`)
  exit(0)
} catch (err) {
  console.error(`${COLOR.fgRed}${err.message}${COLOR.reset}`)
  exit(1)
}