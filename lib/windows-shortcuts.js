var execFile = require('child_process').execFile;

/*
 * options object (also passed by query())
 * target : The path the shortcut points to
 * args : The arguments passed to the target as a string
 * workingDir : The working directory of the target
 * runStyle : State to open the window in: ws.NORMAL (1), ws.MAX (3), or ws.MIN (7)
 * icon : The path to the shortcut icon file
 * iconIndex : An optional index for the image in the icon file
 * hotkey : A numerical hotkey
 * desc : A description
 */

function parseQuery(stdout) {
	// Parses the stdout of a shortcut.exe query into a JS object
	var result = {};
	result.expanded = {};
	stdout.split(/[\r\n]+/)
		.filter(function(line) { return line.indexOf('=') !== -1; })
		.forEach(function(line) {
				var pair = line.split('=', 2),
				key = pair[0],
				value = pair[1];
				if (key === "TargetPath")
					result.target = value;
				else if (key === "TargetPathExpanded")
					result.expanded.target = value;
				else if (key === "Arguments")
					result.args = value;
				else if (key === "ArgumentsExpanded")
					result.expanded.args = value;
				else if (key === "WorkingDirectory")
					result.workingDir = value;
				else if (key === "WorkingDirectoryExpanded")
					result.expanded.workingDir = value;
				else if (key === "RunStyle")
					result.runStyle = +value;
				else if (key === "IconLocation") {
					result.icon = value.split(',')[0];
					result.iconIndex = value.split(',')[1];
				} else if (key === "IconLocationExpanded") {
					result.expanded.icon = value.split(',')[0];
				} else if (key === "HotKey")
					result.hotkey = +value.match(/\d+/)[0];
				else if (key === "Description")
					result.desc = value;
			});
	Object.keys(result.expanded).forEach(function(key) {
		result.expanded[key] = result.expanded[key] || result[key];
	});
	return result;
}

function generateCommand(type, path, options) {
	// Generates a command for shortcut.exe
	var command = [];
	command.push('/A:' + type);
	command.push('/F:"' + path + '"');

	if (options) {
		if (options.target)
			command.push('/T:"' + options.target.replace(/(\^%[^^]*\^%)/g, '"$1"') + '"'); // ^% environment variables can't be inside quotation marks
		if (options.args)
			command.push('/P:"' + options.args.replace('"','\\"','g') + '"');
		if (options.workingDir)
			command.push('/W:"' + options.workingDir + '"');
		if (options.runStyle)
			command.push('/R:' + options.runStyle);
		if (options.icon) {
			command.push('/I:"' + options.icon + '"');
			if (options.iconIndex)
				command.push(',' + options.iconIndex);
		}
		if (options.hotkey)
			command.push('/H:' + options.hotkey);
		if (options.desc)
			command.push('/D:"' + options.desc + '"');
	}
	return command;
}

function isString(x) {
	return Object.prototype.toString.call(x) === "[object String]";
}

exports.query = function(path, callback) {
	var file = __dirname + '/shortcut/shortcut.exe';
	execFile(file, ['/A:Q', '/F:"' + path + '"'],
	     function(error, stdout, stderr) {
		 	var result = parseQuery(stdout);
			callback(error ? stderr || stdout : null, result);
		});
};

exports.create = function(path, optionsOrCallbackOrTarget, callback) {
	var options = isString(optionsOrCallbackOrTarget) ? {target : optionsOrCallbackOrTarget} : optionsOrCallbackOrTarget;
	callback = typeof optionsOrCallbackOrTarget === 'function' ? optionsOrCallbackOrTarget : callback;

	if (path.indexOf('.lnk') === -1) { // Automatically generate shortcut if a .lnk file name is not given
		path = path.replace(/[\\\/]?$/, (path ? "\\" : "") +
			(options && options.target ?
				options.target.match(/[^\\\/]+(?=\..*$)/)[0] + ".lnk" :
	        	"New Shortcut.lnk"));
	}

	var file = __dirname + '/shortcut/shortcut.exe';
	execFile(file, generateCommand('C', path, options),
	     function(error, stdout, stderr) {
		 	if (callback)
				callback(error ? stderr || stdout : null);
		});
};

exports.edit = function(path, options, callback) {
	var file = __dirname + '/shortcut/shortcut.exe';
	execFile(file, generateCommand('E', path, options),
	     function(error, stdout, stderr) {
		 	if (callback)
				callback(error ? stderr || stdout : null);
		});
};

exports.addAppId = function(path, appId, callback) {
	var file = __dirname + '/Win7AppId/Win7AppId.exe';
	var command = [];
	command.push(' "' + path + '"');
	command.push(' "' + appId + '"');
	execFile(file, command, function(error, stdout, stderr) {
		if (callback)
		  callback(error ? stderr || stdout : null);
	});
};

// Shortcut open states
exports.NORMAL = 1;
exports.MAX = 3;
exports.MIN = 7;
