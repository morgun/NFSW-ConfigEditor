/**
 * @author Khavilo "widowmaker" Dmitry <me@widowmaker.kiev.ua>
 * @version 0.1.2
 * @type {HTMLElement}
 */
var	de = document.documentElement,
		win = {w:500, h: 648},
		id = function (id) {
			return document.getElementById(id);
		},
		isset = function (variable) {
			return typeof variable !== 'undefined';
		},
		drag = {x:0, y:0},
		currentConfiguration = {};

var presets = {
	low:{
		texturefilter:"0",
		texturemaxani:"1",
		carenvironmentmapenable:"0",
		carlodlevel:"0",
		enableaero:"0",
		fsaalevel:"0",
		globaldetaillevel:"0",
		maxskidmarks:"0",
		motionblurenable:"0",
		overbrightenable:"0",
		particlesystemenable:"1",
		rainenable:"0",
		roadreflectionenable:"0",
		shaderdetail:"0",
		shadowdetail:"0",
		visualtreatment:"0",
		vsyncon:"0",
		watersimenable:"0"
	},
	highest:{
		texturefilter:"2",
		texturemaxani:"16",
		carenvironmentmapenable:"4",
		carlodlevel:"1",
		enableaero:"1",
		fsaalevel:"8",
		globaldetaillevel:"4",
		maxskidmarks:"2",
		motionblurenable:"1",
		overbrightenable:"1",
		particlesystemenable:"1",
		rainenable:"1",
		roadreflectionenable:"2",
		shaderdetail:"4",
		shadowdetail:"2",
		visualtreatment:"1",
		vsyncon:"1",
		watersimenable:"1"
	}
};

Object.prototype.toSource = function () {
	var result = '{\n';
	for (var i in this)
		if (this.hasOwnProperty(i))
			result += '  ' + i + ': "' + this[i] + '",\r\n';
	result += '}';
	return result.replace(',\r\n}', '\r\n}');
};

window.resizeTo(win.w, win.h);
win = {w:2 * win.w - de.offsetWidth, h:2 * win.h - de.offsetHeight},
		window.resizeTo(win.w, win.h);

/** Enable window dragging **/
document.body.attachEvent('onmousedown', function (e) {
	drag = {x:e.x, y:e.y};
});
document.body.attachEvent('onmousemove', function (e) {
	if (e.button === 1) {
		window.moveTo(window.screenX - drag.x + e.x, window.screenY - drag.y + e.y);
	}
});

var Locator = new ActiveXObject("WbemScripting.SWbemLocator");
var WMI = Locator.ConnectServer('.', "/root/CIMV2");
var FSO = new ActiveXObject("Scripting.FileSystemObject");
var WShell = new ActiveXObject("Wscript.Shell");
var APPDATA = WShell.ExpandEnvironmentStrings("%APPDATA%");

var configFile = APPDATA + '\\Need for Speed World\\Settings\\UserSettings.xml';
var configBackup = APPDATA + '\\Need for Speed World\\Settings\\UserSettings.xml.bak';
var configTagRegex = /<VideoConfig>\s*(.*?)\s*<\/VideoConfig>/gmi;

function ReadFile(filename) {
	var fp = FSO.OpenTextFile(filename, 1, false, 0),
			fileContents = fp.ReadAll();
	fp.Close();
	return fileContents;
}

function WriteFile(filename, fileContents) {
	var fp = FSO.OpenTextFile(filename, 2, true, 0);
	fp.Write(fileContents);
	fp.Close();
}

try {
	var config = ReadFile(configFile);
	if (!FSO.FileExists(configBackup)) {
		FSO.CopyFile(configFile, configBackup)
	}
} catch (e) {
	alert('Config not found.\nGame in not installed or wasn\'t launched yet.');
	self.close();
}

try {
	var template = ReadFile('template.tpl');
} catch (e) {
	alert('Error reading template.');
	self.close();
}

parseConfig();

function parseConfig() {
	var tempConfig = configTagRegex.exec(config.replace(/\r\n/g, ''))[1].match(/<[^\/].*?>[^<]+/g),
			tagParser = /<(\S+).*?>(.*)/,
			temp = [],
			key = '',
			value = '';
	for (var i = 0; i < tempConfig.length; i++) {
		try {
			temp = tagParser.exec(tempConfig[i]);
			key = temp[1];
			value = temp[2];
			switch (key) {
				case 'basetexturefilter':
				case 'basetexturemaxani':
					currentConfiguration[key.replace(/^base/i, '')] = value;
					break;
				case 'basetexturelodbias':
				case 'firsttime':
				case 'forcesm1x':
				case 'roadtexturefilter':
				case 'roadtexturelodbias':
				case 'roadtexturemaxani':
				case 'screenleft':
				case 'screentop':
				case 'size':
					break;
				default:
					currentConfiguration[key] = value;
			}
		} catch (e) {
			alert([tempConfig[i], temp])
		}
	}
}

/** Collection available resolutions **/
var resolutions = {};
var resCollection = WMI.ExecQuery("SELECT * FROM CIM_VideoControllerResolution");
var max = {square:0, resolution:''};
var min = {square:0xFFFFFFF, resolution:''};

var e = new Enumerator(resCollection);
var strProcess = '';
var res = '';
for (; !e.atEnd(); e.moveNext()) {
	var i = e.item();
	if (i.NumberOfColors === '4294967296') {
		res = i.HorizontalResolution + 'x' + i.VerticalResolution;

		if (max.square < (i.HorizontalResolution * i.VerticalResolution)) {
			max.square = (i.HorizontalResolution * i.VerticalResolution);
			max.resolution = res;
		}

		if (isset(resolutions[res]))
			resolutions[res].push(i.RefreshRate);
		else
			resolutions[res] = [i.RefreshRate];
	}
}

var aspect = eval(max.resolution.replace('x', '/'));

for (i in resolutions) {
	if (resolutions.hasOwnProperty(i) && eval(i.replace('x', '/')) === aspect && min.square > (min.square = Math.min(min.square, eval(i.replace('x', '*'))))) {
		presets.low.screenwidth = i.split('x')[0];
		presets.low.screenheight = i.split('x')[1];
	}
}

presets.highest.screenwidth = max.resolution.split('x')[0];
presets.highest.screenheight = max.resolution.split('x')[1];


window.onload = function () {

	/** Loading screen resolutions **/
	var res = id('resolution');
	var rr = id('refresh_rate');
	for (var i in resolutions)
		if (resolutions.hasOwnProperty(i))
			res.options[res.options.length] = new Option(i);

	res.selectedIndex = res.options.length - 1;

	/** Refresh rate dependency **/
	res.attachEvent('onclick', function () {
		rr.options.length = 0;
		var selectedRes = resolutions[res.options[res.selectedIndex].value];
		for (var i in selectedRes)
			if (selectedRes.hasOwnProperty(i))
				rr.options[rr.options.length] = new Option(selectedRes[i] + 'Hz', selectedRes[i]);
	});
	res.click();
	res.blur();

	/** Cancel bubbling for form elements **/
	for (var elementName in {'select':true, 'input':true, 'button':true})
		if (typeof elementName === 'string') {
			var elements = document.getElementsByTagName(elementName);
			for (i = 0; i < elements.length; i++) {
				elements[i].attachEvent('onmousemove', function (e) {
					e.cancelBubble = true;
				});
			}
		}

	loadSettings();
};

function setSelect(select, value) {
	var i = 0;
	while (i < select.options.length) {
		if (select.options[i].value == value) {
			select.selectedIndex = i;
			break;
		}
		i++;
	}
}

function round(value, precision) {
	var mult = 1;
	for(var i = precision; i--;) mult *= 10;
	return Math.round(value * mult) / mult;
}

function loadSettings() {
	/** Writing settings to form **/
	var elements = document.getElementsByTagName('select'), el, val;
	for (var i = 0; i < elements.length; i++) {
		el = elements[i];

		val = el.options[el.selectedIndex].value;
		if (el.name) {
			setSelect(el, currentConfiguration[el.name]);
		} else {
			switch (el.id) {
				case 'resolution':
					setSelect(el, currentConfiguration.screenwidth + 'x' + currentConfiguration.screenheight);
					break;
				case 'aspect':
					/*** Game saves pixel width ***/
					setSelect(el, round(parseFloat(currentConfiguration.pixelaspectratiooverride) * round(currentConfiguration.screenwidth / currentConfiguration.screenheight, 2), 2));
					break;
			}
		}
	}

}

/** Reading settings from form **/
function readSetting() {
	var elements = document.getElementsByTagName('select'), el, val;
	for (var i = 0; i < elements.length; i++) {
		el = elements[i];
		val = el.options[el.selectedIndex].value;
		if (el.name) {
			currentConfiguration[el.name] = val;
		} else {
			switch (el.id) {
				case 'resolution':
					val = val.split('x');
					currentConfiguration.screenwidth = val[0];
					currentConfiguration.screenheight = val[1];
					break;
				case 'aspect':
					/*** Game saves pixel width ***/
					currentConfiguration.pixelaspectratiooverride = round(parseFloat(val) / round(currentConfiguration.screenwidth / currentConfiguration.screenheight, 2), 3);
					break;
			}
		}
	}
}

function save() {
	readSetting();
	var tpl = template + '';
	/** Prevent copying **/
	for (var i in currentConfiguration)
		if (currentConfiguration.hasOwnProperty(i))
			tpl = tpl.replace(new RegExp('{' + i + '}', 'gi'), currentConfiguration[i]);

	WriteFile(configFile, config.replace(/\r\n/g, '<RN>').replace(configTagRegex, '<VideoConfig>\r\n' + tpl + '\r\n    </VideoConfig>').replace(/<RN>/g, '\r\n'));
}


function savePreset() {
	readSetting();
	WriteFile('settings.js', 'var settings = ' + currentConfiguration.toSource());
}

function setPreset(preset) {
	for (var i in presets[preset])
		if (presets[preset].hasOwnProperty(i))
			currentConfiguration[i] = presets[preset][i];

	loadSettings();
}