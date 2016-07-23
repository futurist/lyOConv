var path = require('path')
var fs = require('fs')
var spawn = require('child_process').spawn

var folder = 'D:\\Test'

try {
  var CONFIG = require('../config.json')
  folder = path.join(CONFIG.filesFolder, 'ÎÄ¼þ×ª»»')
} catch(e) {}

var fileStore = {}
var INTERVAL = 300

function getFile (name) {
  return path.join(folder, name)
}

function checkFile (filename) {
  var file = getFile(filename)
  var store = fileStore[filename]
  fs.stat(file, function (err, stats) {
    if (store.stats.size == stats.size) {
      convertFile(file)
    } else {
      store.stats = stats
      setTimeout(function () {
        checkFile(filename)
      }, INTERVAL)
    }
  })
}

fs.watch(folder, {persistent: true}, function (event, filename) {
  // add or delete
  if (event == 'rename') {
    // it's second time rename (delete)
    if (filename in fileStore) {
      delete fileStore[filename]
    } else {
      fileStore[filename] = {stats: {}}
      checkFile(filename)
    }
  }
})

console.log('watching folder for office files', folder)

function convertFile (file) {
  if( /^\~\$/.test(path.basename(file)) ) return
  if (path.extname(file).toLowerCase().match(/x$/i)) {
    console.log('File', file, 'has been added')
    spawn('cscript', ['oconv.vbs', file, path.join(folder, 'ok')])
  } else {
    fs.unlink(file)
  }
}
