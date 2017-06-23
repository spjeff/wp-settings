/**
 * Shrinks promoted links on a page to make them smaller (120px vs the 160px default) and allows for wrapping (defaults to 6 per row).
 *
 * version 0.1.15
 * last updated 05-11-2017
 *
 */

String.prototype.replaceAll = function (search, replacement) {
	var target = this;
	return target.split(search).join(replacement);
};

function addCss(css) {
	// dynamically insert CSS to page header
	var head = document.getElementsByTagName('head')[0];
	var s = document.createElement('style');
	s.setAttribute('type', 'text/css');
	if (s.styleSheet) {
		// IE
		s.styleSheet.cssText = css;
	} else {
		// the world
		s.appendChild(document.createTextNode(css));
	}
	head.appendChild(s);
}

function resizePromotedWebPart(settings, containerId) {
	// set height and width of promoted images
	var wrap = settings.imageCount * (settings.imageSize + 10);
	var size = settings.imageSize;
	var sizePadded = size + 10;
	var topOffset = 0;
	if (size < 140) {
		topOffset = size - 150;
	}
	var css = ' #WebPartWPQ2 .ms-promlink-body {width: ' + wrap + 'px}' +
		' #WebPartWPQ2 .ms-promlink-header {display: none}' +
		' #WebPartWPQ2 div.ms-tileview-tile-titleTextMediumCollapsed {background: rgba(0, 46, 79, 0.6) !important;}' +
		' #WebPartWPQ2 .ms-comm-tCat-tile {width: ' + sizePadded + 'px;height: ' + sizePadded + 'px}' +
		' #WebPartWPQ2 .ms-tileview-tile-content img {width:' + size + 'px !important; height:' + size + 'px !important;top: 0.1px;left: 0.1px} ' +
		' #WebPartWPQ2 div.ms-promlink-body {height: ' + size + 'px;}' +
		' #WebPartWPQ2 div.ms-tileview-tile-root {height: ' + sizePadded + 'px !important;width: ' + sizePadded + 'px !important;}' +
		' #WebPartWPQ2 div.ms-tileview-tile-content, #WebPartWPQ2 div.ms-tileview-tile-detailsBox, #WebPartWPQ2 div.ms-tileview-tile-content > a > div > span {height: ' + size + 'px !important;width: ' + size + 'px !important; }' +
		' #WebPartWPQ2 div.ms-tileview-tile-content > a > div > img {max-width: 100%;width: 100% !important;}' +
		' #WebPartWPQ2 ul.ms-tileview-tile-detailsListMedium {height: ' + size + 'px;padding: 0;}' +
		' #WebPartWPQ2 li.ms-tileview-tile-descriptionMedium {font-size: 11px;line-height: 16px;}' +
		' #WebPartWPQ2 div.ms-tileview-tile-titleTextMediumExpanded, #WebPartWPQ2 div.ms-tileview-tile-titleTextLargeCollapsed, #WebPartWPQ2 div.ms-tileview-tile-titleTextLargeExpanded {padding: 3px;}' +
		' #WebPartWPQ2 div.ms-tileview-tile-titleTextMediumCollapsed {background: none repeat scroll 0 0 rgba(0, 46, 79, 0.6) !important; font-size: 12px;line-height: 16px;min-height: 100px;min-width: ' + size + 'px;padding-left: 3px;position: absolute;top: ' + topOffset + 'px; }' +
		' #WebPartWPQ2 li.ms-tileview-tile-descriptionMedium {font-size: 11px;line-height: 14px;padding: 3px;}' +
		' #WebPartWPQ2 .ms-tileview-tile-detailsBox {background:transparent !important}';
	css = css.replaceAll('#WebPartWPQ2', '#WebPart' + containerId);

	// show web part
	css += ' .ms-promlink-root {display:inline !important}' +
		' .ms-promlink-root {min-width: ' + wrap + 'px}';

	addCss(css);
}

function defaultSettings(settings) {
	// default
	if (!settings) {
		settings = {
			imageCount: 6,
			imageSize: 120
		};
	} else {
		if (typeof(settings) == "string") {
			settings = JSON.parse(settings);
		}
	}
	return settings;
}

// render web part GUI
function ShrinkAndWrapPromotedLinks(settings) {
	settings = defaultSettings(settings);
	// process all web parts
	var promotedLinkBody = document.getElementsByClassName('ms-promlink-body');
	for (var i = 0; i < promotedLinkBody.length; i++) {
		var containerId = promotedLinkBody[i].id.split('_')[1];
		resizePromotedWebPart(settings, containerId);
	}
}

// render settings GUI
function ShrinkAndWrapSettingsDisplay(settings) {
	settings = defaultSettings(settings);
	var html = 'Image Count: <input type="text" id="ShrinkAndWrapImageCount" value="{0}"></br>Image Size: <input type="text" id="ShrinkAndWrapImageSize" value="{1}"><hr/><input type="button" value="Save Settings" onclick="ShrinkAndWrapSettingsSave()"><div id="ShrinkAndWrapSettingsOK" style="background-color: lightgreen; border: 1px; border-color: gray"></div><style>.ms-promlink-root {display:inline !important}</style>';
	html = html.replace('{0}',settings.imageCount);
	html = html.replace('{1}',settings.imageSize);
	var div = document.getElementById('shrinkAndWrapPromotedLinksSettings');
	var node = document.createElement('div');
	node.innerHTML = html;
	div.appendChild(node);
}

// save settings GUI
function ShrinkAndWrapSettingsSave() {
	var settings = {
		imageCount: parseInt(document.getElementById('ShrinkAndWrapImageCount').value),
		imageSize: parseInt(document.getElementById('ShrinkAndWrapImageSize').value)
	};
	wpsWrite(null, JSON.stringify(settings), ngShrinkAndWrapSettingsOK);
}

// saved OK
function ShrinkAndWrapSettingsOK() {
	document.getElementById('ShrinkAndWrapSettingsOK').innerHTML = 'Saved';
}

// web part initialize
function wpInit() {
	// is current page in Design Mode?
	try {
		var inDesignMode = document.forms[MSOWebPartPageFormName].MSOLayouts_InDesignMode.value;
	} catch (ex) {}
	try {
		var wikiInEditMode = document.forms[MSOWebPartPageFormName]._wikiPageMode.value;
	} catch (ex) {}
	if (inDesignMode || wikiInEditMode == "Edit") {
		//in settings mode
		wpsRead(null, ShrinkAndWrapSettingsDisplay);
	} else {
		//in visitor mode
		wpsRead(null, ShrinkAndWrapPromotedLinks);
	}
}
_spBodyOnLoadFunctionNames.push('wpInit');