var path = require('path')
var fs = require('fs')
var iconv = require('iconv-lite')
var spawn = require('child_process').spawn
var mkdirp=require("mkdirp")

var folder = 'D:\\Test'

try {
  var CONFIG = JSON.parse(fs.readFileSync('../config.json', 'utf8'))
  folder = pathJoin(CONFIG.officeFolder, '')
} catch(e) {}

mkdirp(folder)
console.log(111, folder)

var fileStore = {}
var INTERVAL = 300

function pathJoin(a,b){
  return path.join(a,b)
}

function toGBK(str){
  return iconv.decode(str, 'utf8').toString('gbk')
}

function toUTF8(str){
  return iconv.encode(str, 'utf8')
}

function getFile (name) {
  return pathJoin(folder, name)
}

function checkFile (filename) {
  var file = getFile(filename)
  var store = fileStore[filename]
  fs.stat(file, function (err, stats) {
    if(err) return console.log(err)
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
  if( !filename || /^\~\$/.test(filename) ) return
  // add or delete
  if (event == 'rename') {
	filename = toGBK(filename)
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
  if (path.extname(file).toLowerCase().match(/x$/i)) {
    console.log('File', file, 'has been added')
    spawn('cscript', ['oconv.vbs', file, pathJoin(folder, 'ok')])
  } else {
    fs.unlink(file, function(err) {
      if(err) return
    })
  }
}
