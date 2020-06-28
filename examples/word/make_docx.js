var async = require('async')
var officegen = require('../../') // In your code replace it with 'officegen'.

var fs = require('fs')
var path = require('path')

// We'll put the generated document in the tmp folder under the root folder of officegen:
var outDir = path.join(__dirname, '../../tmp/')

// Create the document object:
var docx = officegen({
  type: 'docx',
  orientation: 'portrait',
  pageMargins: { top: 1000, left: 1000, bottom: 1000, right: 1000 },
  columns: 2
})

// Remove this comment if you want to debug Officegen:
// officegen.setVerboseMode ( true )

docx.on('error', function (err) {
  console.log(err)
})

var pObj = docx.createP()

// console.log('pobj here', pObj);
global.my_obj = pObj

pObj.addText('Simple')
pObj.addText('Bold + underline')

var table = [
  [
    {
      val: 'No.',
      opts: {
        cellColWidth: 4261,
        color: '7F7F7F',
        b: true,
        sz: '48',
        fontFamily: 'Avenir Book'
      }
    },
    {
      val: 'Title1',
      opts: {
        targetLink: 'https://www.index.hr/',
        b: true,
        color: '7F7F7F',
        align: 'right'
      }
    }
  ]
]

var tableStyle = {
  tableColWidth: 4261,
  tableSize: 24,
  tableColor: 'ada',
  tableAlign: 'left',
  tableFontFamily: 'Comic Sans MS'
}

pObj = docx.createTable(table, tableStyle)

var out = fs.createWriteStream('example12.docx')

out.on('error', function (err) {
  console.log(err)
})

async.parallel(
  [
    function (done) {
      out.on('close', function () {
        console.log('Finish to create a DOCX file.')
        done(null)
      })
      docx.generate(out)
    }
  ],
  function (err) {
    if (err) {
      console.log('error: ' + err)
    } // Endif.
  }
)
