import xlsx from 'xlsx'
import fs from 'fs'
import yargs from 'yargs'
import * as cheerio from 'cheerio'
import { hideBin } from 'yargs/helpers'
import { exit } from 'process'
import COLOR from '../utils/colors.js'
import { defaultConfig, dirtyConfig } from '../utils/config.js'

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
  .option('dirty', {
    describe: 'Convert XLSX file to XML without cleaning.',
    type: 'boolean',
    default: false
  })
  .option('--ignore-halves', {
    describe: 'Ignore rows that has value only in one column.',
    type: 'boolean',
    default: false
  })
  .help()
  .alias('help', 'h')
  .argv

if (!argv.input) {
  // Fail if no input file is set.
  console.error(`${COLOR.fgRed}Please provide an input file, use --help or -h for help.${COLOR.reset}`)
  exit(0)
}

// Set global variables
const DIRTY_OUTPUT = argv.dirty
const IGNORE_HALVES = argv.ignoreHalves
const INPUT_FILE = argv.input
let OUTPUT_FILE = argv.output || argv.input.replace('.xlsx', '.xml')
let CONFIG = defaultConfig
if (argv.config) {
  // Set user defined config
  CONFIG = JSON.parse(fs.readFileSync(argv.config))
}
if (DIRTY_OUTPUT) {
  // Set the config file to dirty and change output filename.
  CONFIG = dirtyConfig
  OUTPUT_FILE = OUTPUT_FILE.replace('.xml', '-dirty.xml')
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

const fixXHTML = (xml) => {
  const $ = cheerio.load(xml, { xmlMode: true, decodeEntities: false })
  return ($.html())
}

/**
 * Clean and format the column value.
 *
 * @param {string} text value to be cleaned
 * @returns {string}
 */
const cleanAndFormatColumnValue = (text) => {
  if (DIRTY_OUTPUT) {
    return fixXHTML(text)
  }

  try {
    let cleanedText = stripTags(text)
    cleanedText = replaceTags(cleanedText)
    cleanedText = divideBy(cleanedText)

    return fixXHTML(cleanedText)
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

/**
 * Check following rows for content, if there is none then append column value
 * to the previous (current) row.
 *
 * @param {Object} sheet the XLSX sheet object
 * @param {number} rowIndex the row index
 * @param {Array<number>} cells indexes of cells to be checked for content
 * @param {string} content the content of the current row
 * @param {number} parentRowIndex the index of parent row, eg. the last row
 *                                that had content in both columns
 * @returns {string}
 */
const checkNextRowForContent = (sheet, rowIndex, cells = [0, 1], content = '', parentRowIndex) => {
  const emptyCell = sheet[xlsx.utils.encode_cell({ r: rowIndex + 1, c: cells[0] })]
  const contentCell = sheet[xlsx.utils.encode_cell({ r: rowIndex + 1, c: cells[1] })]

  if (!emptyCell &&contentCell) {
    content = ` ${cleanAndFormatColumnValue(contentCell.h)}`
    console.info(`${COLOR.fgYellow}[${rowIndex + 2}]\t${COLOR.reset}Has one cell empty, the other cell: "${COLOR.fgBlue}${content}${COLOR.reset}" was appended to the row #${parentRowIndex}`)

    return checkNextRowForContent(sheet, rowIndex + 1, cells, content, parentRowIndex)
  }

  return content
}

/**
 * Create XML file and save it to the disk.
 * @param {string} fileName the name of the file
 * @param {string} content the content of the file
 * @returns {void}
*/
const storeFile = (fileName, content) => {
  fs.writeFile(fileName, content, (err) => {
    if (err) {
      console.info(`\n${COLOR.bgRed}${COLOR.fgBlack}Error when saving file.${COLOR.reset}`)
      console.error(`${COLOR.fgRed}${err.message}${COLOR.reset}\n`)
      exit(1)
    } else {
      console.info(`\n${COLOR.bgGreen}${COLOR.fgBlack}Successfully created file: ${OUTPUT_FILE}${COLOR.reset}`)
      exit(0)
    }
  })
}

// Read the input file
try {
  console.info(`${COLOR.bgBlue}${COLOR.fgBlack}Starting conversion of file: ${INPUT_FILE}${COLOR.reset}\n`)

  const file = xlsx.readFile(INPUT_FILE)
  const sheet = file.Sheets[file.SheetNames[0]]
  const range = xlsx.utils.decode_range(sheet['!ref'])
  const code = []

  for (let rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
    const firstCell = sheet[xlsx.utils.encode_cell({ r: rowNum, c: 0 })]
    const secondCell = sheet[xlsx.utils.encode_cell({ r: rowNum, c: 1 })]

    if (firstCell && secondCell) {
      const idFirstCell = `${firstCell.v.slice(0, 1).toUpperCase()}-0-${rowNum + 1}`
      const idSecondCell = `${firstCell.v.slice(0, 1).toUpperCase()}-1-${rowNum + 1}`

      // Check for additional content in the following rows in case some of them has one cell empty.
      let additionToFirstCell = ''
      let additionToSecondCell = ''
      if (!IGNORE_HALVES) {
        additionToSecondCell = await checkNextRowForContent(sheet, rowNum, [0, 1], '', rowNum + 1)
        additionToFirstCell = await checkNextRowForContent(sheet, rowNum, [1, 0], '', rowNum + 1)
      }

      // Language definitions if there are any
      const firstColumnLanguage = CONFIG?.language?.length > 0 ? `xml:lang="${CONFIG.language[0]}"` : ''
      const secondColumnLanguage = CONFIG?.language?.length > 1 ? `xml:lang="${CONFIG.language[1]}"` : ''

      // The code snippet
      code.push(`  <${CONFIG.rowTagName} id="${idFirstCell}" corresp="${idSecondCell}" ${firstColumnLanguage}>${cleanAndFormatColumnValue(firstCell.h) + additionToFirstCell}</${CONFIG.rowTagName}>\n  <${CONFIG.rowTagName} id="${idSecondCell}" corresp="${idFirstCell}" ${secondColumnLanguage}>${cleanAndFormatColumnValue(secondCell.h) + additionToSecondCell}</${CONFIG.rowTagName}>\n`)
    } else if (secondCell) {
      // console.log(`${COLOR.dim}[${rowNum + 1}]\tSecond cell is empty.${COLOR.reset}`)
      // Do something when only first cell is empty
    } else if (firstCell) {
      // console.log(`${COLOR.dim}[${rowNum + 1}]\tFirst cell is empty.${COLOR.reset}`)
      // Do something when only second cell is empty
    } else {
      console.warn(`${COLOR.fgRed}[${rowNum + 1}]\tRow is empty${COLOR.reset}`)
    }
  }

  let content = `<?xml version="1.0" encoding="UTF-8" ?>\n<${CONFIG.parentTagName}>\n`
  code.map(line => {
    content += line
  })
  content += `</${CONFIG.parentTagName}>\n`

  storeFile(OUTPUT_FILE, content)
} catch (err) {
  console.error(`${COLOR.fgRed}${err.message}${COLOR.reset}`)
  exit(1)
}
