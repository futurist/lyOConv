var chokidar = require('chokidar')
var path = require('path')
var spawn = require('child_process').spawn

var folder = 'D:\\Test'

// Initialize watcher.
var watcher = chokidar.watch(folder, {
  ignored: /^\~\$/,
  persistent: true,
  depth: 0
})

watcher.on( 'add', function(file) {
  if(path.extname(file).toLowerCase().match(/x$/i) && !/^\~\$/.test(path.basename(file)) ) {
    console.log('File', file, 'has been added')
    spawn('cscript', ['oconv.vbs', file, path.join(folder, 'ok')])
  }
})
