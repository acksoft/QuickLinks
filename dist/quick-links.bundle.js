define("13f86250-30e1-4250-b2d2-f7651a9dd6a5_0.0.1", ["@microsoft/sp-core-library","@microsoft/sp-webpart-base","@microsoft/sp-lodash-subset","quickLinksStrings"], function(__WEBPACK_EXTERNAL_MODULE_1__, __WEBPACK_EXTERNAL_MODULE_2__, __WEBPACK_EXTERNAL_MODULE_3__, __WEBPACK_EXTERNAL_MODULE_9__) { return /******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId])
/******/ 			return installedModules[moduleId].exports;
/******/
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			exports: {},
/******/ 			id: moduleId,
/******/ 			loaded: false
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.loaded = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(0);
/******/ })
/************************************************************************/
/******/ ([
/* 0 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	var __extends = (this && this.__extends) || function (d, b) {
	    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
	    function __() { this.constructor = d; }
	    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
	};
	var sp_core_library_1 = __webpack_require__(1);
	var sp_webpart_base_1 = __webpack_require__(2);
	var sp_lodash_subset_1 = __webpack_require__(3);
	var QuickLinks_module_scss_1 = __webpack_require__(4);
	var strings = __webpack_require__(9);
	__webpack_require__(10);
	var sp_pnp_js_1 = __webpack_require__(11);
	var QuickLinksWebPart = (function (_super) {
	    __extends(QuickLinksWebPart, _super);
	    function QuickLinksWebPart() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    QuickLinksWebPart.prototype.render = function () {
	        if (sp_core_library_1.Environment.type === sp_core_library_1.EnvironmentType.Local) {
	            // this.domElement.innerHTML = this.getLinksHTML();
	            this.domElement.innerHTML = "\n      <div class=\"" + QuickLinks_module_scss_1.default.quickLinks + "\">\n        <div class=\"" + QuickLinks_module_scss_1.default.container + "\">\n          <div class=\"ms-Grid-row ms-bgColor-themeDark ms-fontColor-white " + QuickLinks_module_scss_1.default.row + "\">\n            <div class=\"ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1\">\n              <span class=\"ms-font-xl ms-fontColor-white\">Quick Links (Running Locally)</span>\n              <p class=\"ms-font-l ms-fontColor-white\">" + sp_lodash_subset_1.escape(this.properties.LinkSet) + "</p>\n            </div>\n          </div>\n        </div>\n      </div>";
	        }
	        else if (sp_core_library_1.Environment.type === sp_core_library_1.EnvironmentType.SharePoint ||
	            sp_core_library_1.Environment.type === sp_core_library_1.EnvironmentType.ClassicSharePoint) {
	            this.domElement.innerHTML = "\n      <div class=\"" + QuickLinks_module_scss_1.default.quickLinks + "\">\n        <div class=\"" + QuickLinks_module_scss_1.default.container + "\">\n          <div class=\"ms-Grid-row ms-bgColor-themeDark ms-fontColor-white " + QuickLinks_module_scss_1.default.row + "\">\n            <div class=\"ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1\">\n              <span class=\"ms-font-xl ms-fontColor-white\">Quick Links (Running in SPO)</span>\n              <p class=\"ms-font-l ms-fontColor-white\">" + sp_lodash_subset_1.escape(this.properties.LinkSet) + " Links</p>\n              <div id=\"qlinks\" />\n            </div>\n          </div>\n        </div>\n      </div>";
	            $('#qlinks', this.domElement).html(this.getLinkData());
	        }
	    };
	    Object.defineProperty(QuickLinksWebPart.prototype, "dataVersion", {
	        get: function () {
	            return sp_core_library_1.Version.parse('1.0');
	        },
	        enumerable: true,
	        configurable: true
	    });
	    QuickLinksWebPart.prototype.getLinkData = function () {
	        var linkData = '';
	        var linkFilter = "LinkSet eq '" + this.properties.LinkSet + "' and Visible eq 1";
	        sp_pnp_js_1.default.sp.web.lists.getByTitle('QuickLinks').items.filter(linkFilter).orderBy("Sequence").get().then(function (r) {
	            var row = '';
	            // start the unordered list
	            linkData = "<ul>";
	            // loop through the returned items to create the links
	            r.array.forEach(function (item) {
	                row += "<li><a href='";
	                row += item.URL.Url;
	                row += "'";
	                // if external, open in new window/tab
	                if (item.External) {
	                    row += " target='_blank'";
	                }
	                // if there are comments, turn them into a tooltip
	                if (item.Comments.Length > 0) {
	                    row += " title='";
	                    row += item.Comments;
	                    row += "'";
	                }
	                //set the link text
	                row += ">";
	                row += item.URL.Description;
	                // if extnernal, add the indicator icon
	                if (item.External) {
	                    row += " <i class='ms-Icon ms-Icon--OpenInNewWindow' aria-hidden='true'></i>";
	                }
	                // close the link
	                row += "</a></li>";
	                // add the link to the list of links
	                linkData += row;
	                row = '';
	            });
	            // close the link list
	            linkData += "</ul>";
	        });
	        return linkData;
	    };
	    QuickLinksWebPart.prototype.getPropertyPaneConfiguration = function () {
	        return {
	            pages: [
	                {
	                    header: {
	                        description: strings.PropertyPaneDescription
	                    },
	                    groups: [
	                        {
	                            groupName: strings.BasicGroupName,
	                            groupFields: [
	                                sp_webpart_base_1.PropertyPaneTextField('LinkSet', {
	                                    label: strings.DescriptionFieldLabel
	                                })
	                            ]
	                        }
	                    ]
	                }
	            ]
	        };
	    };
	    return QuickLinksWebPart;
	}(sp_webpart_base_1.BaseClientSideWebPart));
	Object.defineProperty(exports, "__esModule", { value: true });
	exports.default = QuickLinksWebPart;



/***/ },
/* 1 */
/***/ function(module, exports) {

	module.exports = __WEBPACK_EXTERNAL_MODULE_1__;

/***/ },
/* 2 */
/***/ function(module, exports) {

	module.exports = __WEBPACK_EXTERNAL_MODULE_2__;

/***/ },
/* 3 */
/***/ function(module, exports) {

	module.exports = __WEBPACK_EXTERNAL_MODULE_3__;

/***/ },
/* 4 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	/* tslint:disable */
	__webpack_require__(5);
	var styles = {
	    quickLinks: 'quickLinks_1c4dc9e3',
	    container: 'container_1c4dc9e3',
	    row: 'row_1c4dc9e3',
	    listItem: 'listItem_1c4dc9e3',
	    button: 'button_1c4dc9e3',
	    label: 'label_1c4dc9e3',
	};
	Object.defineProperty(exports, "__esModule", { value: true });
	exports.default = styles;
	/* tslint:enable */ 
	


/***/ },
/* 5 */
/***/ function(module, exports, __webpack_require__) {

	var content = __webpack_require__(6);
	var loader = __webpack_require__(8);
	
	if(typeof content === "string") content = [[module.id, content]];
	
	// add the styles to the DOM
	for (var i = 0; i < content.length; i++) loader.loadStyles(content[i][1]);
	
	if(content.locals) module.exports = content.locals;

/***/ },
/* 6 */
/***/ function(module, exports, __webpack_require__) {

	exports = module.exports = __webpack_require__(7)();
	// imports
	
	
	// module
	exports.push([module.id, ".quickLinks_1c4dc9e3 .container_1c4dc9e3{max-width:700px;margin:0 auto;box-shadow:0 2px 4px 0 rgba(0,0,0,.2),0 25px 50px 0 rgba(0,0,0,.1)}.quickLinks_1c4dc9e3 .row_1c4dc9e3{padding:20px}.quickLinks_1c4dc9e3 .listItem_1c4dc9e3{max-width:715px;margin:5px auto 5px auto;box-shadow:0 0 4px 0 rgba(0,0,0,.2),0 25px 50px 0 rgba(0,0,0,.1)}.quickLinks_1c4dc9e3 .button_1c4dc9e3{text-decoration:none;height:32px;min-width:80px;background-color:#78BC20;border-color:#78BC20;color:#fff;outline:transparent;position:relative;font-family:\"Segoe UI WestEuropean\",\"Segoe UI\",-apple-system,BlinkMacSystemFont,Roboto,\"Helvetica Neue\",sans-serif;-webkit-font-smoothing:antialiased;font-size:14px;font-weight:400;border-width:0;text-align:center;cursor:pointer;display:inline-block;padding:0 16px}.quickLinks_1c4dc9e3 .button_1c4dc9e3 .label_1c4dc9e3{font-weight:600;font-size:14px;height:32px;line-height:32px;margin:0 4px;vertical-align:top;display:inline-block}", ""]);
	
	// exports


/***/ },
/* 7 */
/***/ function(module, exports) {

	/*
		MIT License http://www.opensource.org/licenses/mit-license.php
		Author Tobias Koppers @sokra
	*/
	// css base code, injected by the css-loader
	module.exports = function() {
		var list = [];
	
		// return the list of modules as css string
		list.toString = function toString() {
			var result = [];
			for(var i = 0; i < this.length; i++) {
				var item = this[i];
				if(item[2]) {
					result.push("@media " + item[2] + "{" + item[1] + "}");
				} else {
					result.push(item[1]);
				}
			}
			return result.join("");
		};
	
		// import a list of modules into the list
		list.i = function(modules, mediaQuery) {
			if(typeof modules === "string")
				modules = [[null, modules, ""]];
			var alreadyImportedModules = {};
			for(var i = 0; i < this.length; i++) {
				var id = this[i][0];
				if(typeof id === "number")
					alreadyImportedModules[id] = true;
			}
			for(i = 0; i < modules.length; i++) {
				var item = modules[i];
				// skip already imported module
				// this implementation is not 100% perfect for weird media query combinations
				//  when a module is imported multiple times with different media queries.
				//  I hope this will never occur (Hey this way we have smaller bundles)
				if(typeof item[0] !== "number" || !alreadyImportedModules[item[0]]) {
					if(mediaQuery && !item[2]) {
						item[2] = mediaQuery;
					} else if(mediaQuery) {
						item[2] = "(" + item[2] + ") and (" + mediaQuery + ")";
					}
					list.push(item);
				}
			}
		};
		return list;
	};


/***/ },
/* 8 */
/***/ function(module, exports) {

	/* WEBPACK VAR INJECTION */(function(global) {/**
	 * An IThemingInstruction can specify a rawString to be preserved or a theme slot and a default value
	 * to use if that slot is not specified by the theme.
	 */
	"use strict";
	// IE needs to inject styles using cssText. However, we need to evaluate this lazily, so this
	// value will initialize as undefined, and later will be set once on first loadStyles injection.
	var _injectStylesWithCssText;
	// Store the theming state in __themeState__ global scope for reuse in the case of duplicate
	// load-themed-styles hosted on the page.
	var _root = (typeof window === 'undefined') ? global : window; // tslint:disable-line:no-any
	var _themeState = _root.__themeState__ = _root.__themeState__ || {
	    theme: undefined,
	    lastStyleElement: undefined,
	    registeredStyles: []
	};
	/**
	 * Matches theming tokens. For example, "[theme: themeSlotName, default: #FFF]" (including the quotes).
	 */
	/* tslint:disable: max-line-length */
	var _themeTokenRegex = /[\'\"]\[theme:\s*(\w+)\s*(?:\,\s*default:\s*([\\"\']?[\.\,\(\)\#\-\s\w]*[\.\,\(\)\#\-\w][\"\']?))?\s*\][\'\"]/g;
	/* tslint:enable: max-line-length */
	/** Maximum style text length, for supporting IE style restrictions. */
	var MAX_STYLE_CONTENT_SIZE = 10000;
	/**
	 * Loads a set of style text. If it is registered too early, we will register it when the window.load
	 * event is fired.
	 * @param {string | ThemableArray} styles Themable style text to register.
	 */
	function loadStyles(styles) {
	    var styleParts = Array.isArray(styles) ? styles : splitStyles(styles);
	    if (_injectStylesWithCssText === undefined) {
	        _injectStylesWithCssText = shouldUseCssText();
	    }
	    applyThemableStyles(styleParts);
	}
	exports.loadStyles = loadStyles;
	/**
	 * Allows for customizable loadStyles logic. e.g. for server side rendering application
	 * @param {(styles: string) => void} a loadStyles callback that gets called when styles are loaded or reloaded
	 */
	function configureLoadStyles(callback) {
	    _themeState.loadStyles = callback;
	}
	exports.configureLoadStyles = configureLoadStyles;
	/**
	 * Loads a set of style text. If it is registered too early, we will register it when the window.load event
	 * is fired.
	 * @param {string} styleText Style to register.
	 * @param {IStyleRecord} styleRecord Existing style record to re-apply.
	 */
	function applyThemableStyles(stylesArray, styleRecord) {
	    if (_themeState.loadStyles) {
	        var styles = resolveThemableArray(stylesArray);
	        _themeState.loadStyles(styles);
	    }
	    else {
	        _injectStylesWithCssText ?
	            registerStylesIE(stylesArray, styleRecord) :
	            registerStyles(stylesArray, styleRecord);
	    }
	}
	/**
	 * Registers a set theme tokens to find and replace. If styles were already registered, they will be
	 * replaced.
	 * @param {theme} theme JSON object of theme tokens to values.
	 */
	function loadTheme(theme) {
	    _themeState.theme = theme;
	    // reload styles.
	    reloadStyles();
	}
	exports.loadTheme = loadTheme;
	/**
	 * Reloads styles.
	 */
	function reloadStyles() {
	    if (_themeState.theme) {
	        for (var _i = 0, _a = _themeState.registeredStyles; _i < _a.length; _i++) {
	            var styleRecord = _a[_i];
	            applyThemableStyles(styleRecord.themableStyle, styleRecord);
	        }
	    }
	}
	/**
	 * Find theme tokens and replaces them with provided theme values.
	 * @param {string} styles Tokenized styles to fix.
	 */
	function detokenize(styles) {
	    if (styles) {
	        styles = resolveThemableArray(splitStyles(styles));
	    }
	    return styles;
	}
	exports.detokenize = detokenize;
	/**
	 * Resolves ThemingInstruction objects in an array and joins the result into a string.
	 * @param {ThemableArray} splitStyleArray ThemableArray to resolve and join.
	 */
	function resolveThemableArray(splitStyleArray) {
	    var theme = _themeState.theme;
	    var resolvedCss;
	    if (splitStyleArray) {
	        // Resolve the array of theming instructions to an array of strings.
	        // Then join the array to produce the final CSS string.
	        var resolvedArray = splitStyleArray.map(function (currentValue) {
	            var themeSlot = currentValue.theme;
	            if (themeSlot) {
	                // A theming annotation. Resolve it.
	                var themedValue = theme ? theme[themeSlot] : undefined;
	                var defaultValue = currentValue.defaultValue;
	                // Warn to console if we hit an unthemed value even when themes are provided.
	                // Allow the themedValue to be undefined to explicitly request the default value.
	                if (theme && !themedValue && console && !(themeSlot in theme)) {
	                    /* tslint:disable: max-line-length */
	                    console.warn("Theming value not provided for \"" + themeSlot + "\". Falling back to \"" + (defaultValue || 'inherit') + "\".");
	                }
	                return themedValue || defaultValue || 'inherit';
	            }
	            else {
	                // A non-themable string. Preserve it.
	                return currentValue.rawString;
	            }
	        });
	        resolvedCss = resolvedArray.join('');
	    }
	    return resolvedCss;
	}
	/**
	 * Split tokenized CSS into an array of strings and theme specification objects
	 * @param {string} styles Tokenized styles to split.
	 */
	function splitStyles(styles) {
	    var result = [];
	    if (styles) {
	        var pos = 0; // Current position in styles.
	        var tokenMatch = void 0;
	        while (tokenMatch = _themeTokenRegex.exec(styles)) {
	            var matchIndex = tokenMatch.index;
	            if (matchIndex > pos) {
	                result.push({
	                    rawString: styles.substring(pos, matchIndex)
	                });
	            }
	            result.push({
	                theme: tokenMatch[1],
	                defaultValue: tokenMatch[2] // May be undefined
	            });
	            // index of the first character after the current match
	            pos = _themeTokenRegex.lastIndex;
	        }
	        // Push the rest of the string after the last match.
	        result.push({
	            rawString: styles.substring(pos)
	        });
	    }
	    return result;
	}
	exports.splitStyles = splitStyles;
	/**
	 * Registers a set of style text. If it is registered too early, we will register it when the
	 * window.load event is fired.
	 * @param {ThemableArray} styleArray Array of IThemingInstruction objects to register.
	 * @param {IStyleRecord} styleRecord May specify a style Element to update.
	 */
	function registerStyles(styleArray, styleRecord) {
	    var head = document.getElementsByTagName('head')[0];
	    var styleElement = document.createElement('style');
	    styleElement.type = 'text/css';
	    styleElement.appendChild(document.createTextNode(resolveThemableArray(styleArray)));
	    if (styleRecord) {
	        head.replaceChild(styleElement, styleRecord.styleElement);
	        styleRecord.styleElement = styleElement;
	    }
	    else {
	        head.appendChild(styleElement);
	    }
	    if (!styleRecord) {
	        _themeState.registeredStyles.push({
	            styleElement: styleElement,
	            themableStyle: styleArray
	        });
	    }
	}
	/**
	 * Registers a set of style text, for IE 9 and below, which has a ~30 style element limit so we need
	 * to register slightly differently.
	 * @param {ThemableArray} styleArray Array of IThemingInstruction objects to register.
	 * @param {IStyleRecord} styleRecord May specify a style Element to update.
	 */
	function registerStylesIE(styleArray, styleRecord) {
	    var head = document.getElementsByTagName('head')[0];
	    var lastStyleElement = _themeState.lastStyleElement, registeredStyles = _themeState.registeredStyles;
	    var stylesheet = lastStyleElement ? lastStyleElement.styleSheet : undefined;
	    var lastStyleContent = stylesheet ? stylesheet.cssText : '';
	    var lastRegisteredStyle = registeredStyles[registeredStyles.length - 1];
	    var resolvedStyleText = resolveThemableArray(styleArray);
	    if (!lastStyleElement || (lastStyleContent.length + resolvedStyleText.length) > MAX_STYLE_CONTENT_SIZE) {
	        lastStyleElement = document.createElement('style');
	        lastStyleElement.type = 'text/css';
	        if (styleRecord) {
	            head.replaceChild(lastStyleElement, styleRecord.styleElement);
	            styleRecord.styleElement = lastStyleElement;
	        }
	        else {
	            head.appendChild(lastStyleElement);
	        }
	        if (!styleRecord) {
	            lastRegisteredStyle = {
	                styleElement: lastStyleElement,
	                themableStyle: styleArray
	            };
	            registeredStyles.push(lastRegisteredStyle);
	        }
	    }
	    lastStyleElement.styleSheet.cssText += detokenize(resolvedStyleText);
	    Array.prototype.push.apply(lastRegisteredStyle.themableStyle, styleArray); // concat in-place
	    // Preserve the theme state.
	    _themeState.lastStyleElement = lastStyleElement;
	}
	/**
	 * Checks to see if styleSheet exists as a property off of a style element.
	 * This will determine if style registration should be done via cssText (<= IE9) or not
	 */
	function shouldUseCssText() {
	    var useCSSText = false;
	    if (typeof document !== 'undefined') {
	        var emptyStyle = document.createElement('style');
	        emptyStyle.type = 'text/css';
	        useCSSText = !!emptyStyle.styleSheet;
	    }
	    return useCSSText;
	}
	
	
	/* WEBPACK VAR INJECTION */}.call(exports, (function() { return this; }())))

/***/ },
/* 9 */
/***/ function(module, exports) {

	module.exports = __WEBPACK_EXTERNAL_MODULE_9__;

/***/ },
/* 10 */
/***/ function(module, exports, __webpack_require__) {

	var __WEBPACK_AMD_DEFINE_ARRAY__, __WEBPACK_AMD_DEFINE_RESULT__;/*!
	 * jQuery JavaScript Library v3.2.1
	 * https://jquery.com/
	 *
	 * Includes Sizzle.js
	 * https://sizzlejs.com/
	 *
	 * Copyright JS Foundation and other contributors
	 * Released under the MIT license
	 * https://jquery.org/license
	 *
	 * Date: 2017-03-20T18:59Z
	 */
	( function( global, factory ) {
	
		"use strict";
	
		if ( typeof module === "object" && typeof module.exports === "object" ) {
	
			// For CommonJS and CommonJS-like environments where a proper `window`
			// is present, execute the factory and get jQuery.
			// For environments that do not have a `window` with a `document`
			// (such as Node.js), expose a factory as module.exports.
			// This accentuates the need for the creation of a real `window`.
			// e.g. var jQuery = require("jquery")(window);
			// See ticket #14549 for more info.
			module.exports = global.document ?
				factory( global, true ) :
				function( w ) {
					if ( !w.document ) {
						throw new Error( "jQuery requires a window with a document" );
					}
					return factory( w );
				};
		} else {
			factory( global );
		}
	
	// Pass this if window is not defined yet
	} )( typeof window !== "undefined" ? window : this, function( window, noGlobal ) {
	
	// Edge <= 12 - 13+, Firefox <=18 - 45+, IE 10 - 11, Safari 5.1 - 9+, iOS 6 - 9.1
	// throw exceptions when non-strict code (e.g., ASP.NET 4.5) accesses strict mode
	// arguments.callee.caller (trac-13335). But as of jQuery 3.0 (2016), strict mode should be common
	// enough that all such attempts are guarded in a try block.
	"use strict";
	
	var arr = [];
	
	var document = window.document;
	
	var getProto = Object.getPrototypeOf;
	
	var slice = arr.slice;
	
	var concat = arr.concat;
	
	var push = arr.push;
	
	var indexOf = arr.indexOf;
	
	var class2type = {};
	
	var toString = class2type.toString;
	
	var hasOwn = class2type.hasOwnProperty;
	
	var fnToString = hasOwn.toString;
	
	var ObjectFunctionString = fnToString.call( Object );
	
	var support = {};
	
	
	
		function DOMEval( code, doc ) {
			doc = doc || document;
	
			var script = doc.createElement( "script" );
	
			script.text = code;
			doc.head.appendChild( script ).parentNode.removeChild( script );
		}
	/* global Symbol */
	// Defining this global in .eslintrc.json would create a danger of using the global
	// unguarded in another place, it seems safer to define global only for this module
	
	
	
	var
		version = "3.2.1",
	
		// Define a local copy of jQuery
		jQuery = function( selector, context ) {
	
			// The jQuery object is actually just the init constructor 'enhanced'
			// Need init if jQuery is called (just allow error to be thrown if not included)
			return new jQuery.fn.init( selector, context );
		},
	
		// Support: Android <=4.0 only
		// Make sure we trim BOM and NBSP
		rtrim = /^[\s\uFEFF\xA0]+|[\s\uFEFF\xA0]+$/g,
	
		// Matches dashed string for camelizing
		rmsPrefix = /^-ms-/,
		rdashAlpha = /-([a-z])/g,
	
		// Used by jQuery.camelCase as callback to replace()
		fcamelCase = function( all, letter ) {
			return letter.toUpperCase();
		};
	
	jQuery.fn = jQuery.prototype = {
	
		// The current version of jQuery being used
		jquery: version,
	
		constructor: jQuery,
	
		// The default length of a jQuery object is 0
		length: 0,
	
		toArray: function() {
			return slice.call( this );
		},
	
		// Get the Nth element in the matched element set OR
		// Get the whole matched element set as a clean array
		get: function( num ) {
	
			// Return all the elements in a clean array
			if ( num == null ) {
				return slice.call( this );
			}
	
			// Return just the one element from the set
			return num < 0 ? this[ num + this.length ] : this[ num ];
		},
	
		// Take an array of elements and push it onto the stack
		// (returning the new matched element set)
		pushStack: function( elems ) {
	
			// Build a new jQuery matched element set
			var ret = jQuery.merge( this.constructor(), elems );
	
			// Add the old object onto the stack (as a reference)
			ret.prevObject = this;
	
			// Return the newly-formed element set
			return ret;
		},
	
		// Execute a callback for every element in the matched set.
		each: function( callback ) {
			return jQuery.each( this, callback );
		},
	
		map: function( callback ) {
			return this.pushStack( jQuery.map( this, function( elem, i ) {
				return callback.call( elem, i, elem );
			} ) );
		},
	
		slice: function() {
			return this.pushStack( slice.apply( this, arguments ) );
		},
	
		first: function() {
			return this.eq( 0 );
		},
	
		last: function() {
			return this.eq( -1 );
		},
	
		eq: function( i ) {
			var len = this.length,
				j = +i + ( i < 0 ? len : 0 );
			return this.pushStack( j >= 0 && j < len ? [ this[ j ] ] : [] );
		},
	
		end: function() {
			return this.prevObject || this.constructor();
		},
	
		// For internal use only.
		// Behaves like an Array's method, not like a jQuery method.
		push: push,
		sort: arr.sort,
		splice: arr.splice
	};
	
	jQuery.extend = jQuery.fn.extend = function() {
		var options, name, src, copy, copyIsArray, clone,
			target = arguments[ 0 ] || {},
			i = 1,
			length = arguments.length,
			deep = false;
	
		// Handle a deep copy situation
		if ( typeof target === "boolean" ) {
			deep = target;
	
			// Skip the boolean and the target
			target = arguments[ i ] || {};
			i++;
		}
	
		// Handle case when target is a string or something (possible in deep copy)
		if ( typeof target !== "object" && !jQuery.isFunction( target ) ) {
			target = {};
		}
	
		// Extend jQuery itself if only one argument is passed
		if ( i === length ) {
			target = this;
			i--;
		}
	
		for ( ; i < length; i++ ) {
	
			// Only deal with non-null/undefined values
			if ( ( options = arguments[ i ] ) != null ) {
	
				// Extend the base object
				for ( name in options ) {
					src = target[ name ];
					copy = options[ name ];
	
					// Prevent never-ending loop
					if ( target === copy ) {
						continue;
					}
	
					// Recurse if we're merging plain objects or arrays
					if ( deep && copy && ( jQuery.isPlainObject( copy ) ||
						( copyIsArray = Array.isArray( copy ) ) ) ) {
	
						if ( copyIsArray ) {
							copyIsArray = false;
							clone = src && Array.isArray( src ) ? src : [];
	
						} else {
							clone = src && jQuery.isPlainObject( src ) ? src : {};
						}
	
						// Never move original objects, clone them
						target[ name ] = jQuery.extend( deep, clone, copy );
	
					// Don't bring in undefined values
					} else if ( copy !== undefined ) {
						target[ name ] = copy;
					}
				}
			}
		}
	
		// Return the modified object
		return target;
	};
	
	jQuery.extend( {
	
		// Unique for each copy of jQuery on the page
		expando: "jQuery" + ( version + Math.random() ).replace( /\D/g, "" ),
	
		// Assume jQuery is ready without the ready module
		isReady: true,
	
		error: function( msg ) {
			throw new Error( msg );
		},
	
		noop: function() {},
	
		isFunction: function( obj ) {
			return jQuery.type( obj ) === "function";
		},
	
		isWindow: function( obj ) {
			return obj != null && obj === obj.window;
		},
	
		isNumeric: function( obj ) {
	
			// As of jQuery 3.0, isNumeric is limited to
			// strings and numbers (primitives or objects)
			// that can be coerced to finite numbers (gh-2662)
			var type = jQuery.type( obj );
			return ( type === "number" || type === "string" ) &&
	
				// parseFloat NaNs numeric-cast false positives ("")
				// ...but misinterprets leading-number strings, particularly hex literals ("0x...")
				// subtraction forces infinities to NaN
				!isNaN( obj - parseFloat( obj ) );
		},
	
		isPlainObject: function( obj ) {
			var proto, Ctor;
	
			// Detect obvious negatives
			// Use toString instead of jQuery.type to catch host objects
			if ( !obj || toString.call( obj ) !== "[object Object]" ) {
				return false;
			}
	
			proto = getProto( obj );
	
			// Objects with no prototype (e.g., `Object.create( null )`) are plain
			if ( !proto ) {
				return true;
			}
	
			// Objects with prototype are plain iff they were constructed by a global Object function
			Ctor = hasOwn.call( proto, "constructor" ) && proto.constructor;
			return typeof Ctor === "function" && fnToString.call( Ctor ) === ObjectFunctionString;
		},
	
		isEmptyObject: function( obj ) {
	
			/* eslint-disable no-unused-vars */
			// See https://github.com/eslint/eslint/issues/6125
			var name;
	
			for ( name in obj ) {
				return false;
			}
			return true;
		},
	
		type: function( obj ) {
			if ( obj == null ) {
				return obj + "";
			}
	
			// Support: Android <=2.3 only (functionish RegExp)
			return typeof obj === "object" || typeof obj === "function" ?
				class2type[ toString.call( obj ) ] || "object" :
				typeof obj;
		},
	
		// Evaluates a script in a global context
		globalEval: function( code ) {
			DOMEval( code );
		},
	
		// Convert dashed to camelCase; used by the css and data modules
		// Support: IE <=9 - 11, Edge 12 - 13
		// Microsoft forgot to hump their vendor prefix (#9572)
		camelCase: function( string ) {
			return string.replace( rmsPrefix, "ms-" ).replace( rdashAlpha, fcamelCase );
		},
	
		each: function( obj, callback ) {
			var length, i = 0;
	
			if ( isArrayLike( obj ) ) {
				length = obj.length;
				for ( ; i < length; i++ ) {
					if ( callback.call( obj[ i ], i, obj[ i ] ) === false ) {
						break;
					}
				}
			} else {
				for ( i in obj ) {
					if ( callback.call( obj[ i ], i, obj[ i ] ) === false ) {
						break;
					}
				}
			}
	
			return obj;
		},
	
		// Support: Android <=4.0 only
		trim: function( text ) {
			return text == null ?
				"" :
				( text + "" ).replace( rtrim, "" );
		},
	
		// results is for internal usage only
		makeArray: function( arr, results ) {
			var ret = results || [];
	
			if ( arr != null ) {
				if ( isArrayLike( Object( arr ) ) ) {
					jQuery.merge( ret,
						typeof arr === "string" ?
						[ arr ] : arr
					);
				} else {
					push.call( ret, arr );
				}
			}
	
			return ret;
		},
	
		inArray: function( elem, arr, i ) {
			return arr == null ? -1 : indexOf.call( arr, elem, i );
		},
	
		// Support: Android <=4.0 only, PhantomJS 1 only
		// push.apply(_, arraylike) throws on ancient WebKit
		merge: function( first, second ) {
			var len = +second.length,
				j = 0,
				i = first.length;
	
			for ( ; j < len; j++ ) {
				first[ i++ ] = second[ j ];
			}
	
			first.length = i;
	
			return first;
		},
	
		grep: function( elems, callback, invert ) {
			var callbackInverse,
				matches = [],
				i = 0,
				length = elems.length,
				callbackExpect = !invert;
	
			// Go through the array, only saving the items
			// that pass the validator function
			for ( ; i < length; i++ ) {
				callbackInverse = !callback( elems[ i ], i );
				if ( callbackInverse !== callbackExpect ) {
					matches.push( elems[ i ] );
				}
			}
	
			return matches;
		},
	
		// arg is for internal usage only
		map: function( elems, callback, arg ) {
			var length, value,
				i = 0,
				ret = [];
	
			// Go through the array, translating each of the items to their new values
			if ( isArrayLike( elems ) ) {
				length = elems.length;
				for ( ; i < length; i++ ) {
					value = callback( elems[ i ], i, arg );
	
					if ( value != null ) {
						ret.push( value );
					}
				}
	
			// Go through every key on the object,
			} else {
				for ( i in elems ) {
					value = callback( elems[ i ], i, arg );
	
					if ( value != null ) {
						ret.push( value );
					}
				}
			}
	
			// Flatten any nested arrays
			return concat.apply( [], ret );
		},
	
		// A global GUID counter for objects
		guid: 1,
	
		// Bind a function to a context, optionally partially applying any
		// arguments.
		proxy: function( fn, context ) {
			var tmp, args, proxy;
	
			if ( typeof context === "string" ) {
				tmp = fn[ context ];
				context = fn;
				fn = tmp;
			}
	
			// Quick check to determine if target is callable, in the spec
			// this throws a TypeError, but we will just return undefined.
			if ( !jQuery.isFunction( fn ) ) {
				return undefined;
			}
	
			// Simulated bind
			args = slice.call( arguments, 2 );
			proxy = function() {
				return fn.apply( context || this, args.concat( slice.call( arguments ) ) );
			};
	
			// Set the guid of unique handler to the same of original handler, so it can be removed
			proxy.guid = fn.guid = fn.guid || jQuery.guid++;
	
			return proxy;
		},
	
		now: Date.now,
	
		// jQuery.support is not used in Core but other projects attach their
		// properties to it so it needs to exist.
		support: support
	} );
	
	if ( typeof Symbol === "function" ) {
		jQuery.fn[ Symbol.iterator ] = arr[ Symbol.iterator ];
	}
	
	// Populate the class2type map
	jQuery.each( "Boolean Number String Function Array Date RegExp Object Error Symbol".split( " " ),
	function( i, name ) {
		class2type[ "[object " + name + "]" ] = name.toLowerCase();
	} );
	
	function isArrayLike( obj ) {
	
		// Support: real iOS 8.2 only (not reproducible in simulator)
		// `in` check used to prevent JIT error (gh-2145)
		// hasOwn isn't used here due to false negatives
		// regarding Nodelist length in IE
		var length = !!obj && "length" in obj && obj.length,
			type = jQuery.type( obj );
	
		if ( type === "function" || jQuery.isWindow( obj ) ) {
			return false;
		}
	
		return type === "array" || length === 0 ||
			typeof length === "number" && length > 0 && ( length - 1 ) in obj;
	}
	var Sizzle =
	/*!
	 * Sizzle CSS Selector Engine v2.3.3
	 * https://sizzlejs.com/
	 *
	 * Copyright jQuery Foundation and other contributors
	 * Released under the MIT license
	 * http://jquery.org/license
	 *
	 * Date: 2016-08-08
	 */
	(function( window ) {
	
	var i,
		support,
		Expr,
		getText,
		isXML,
		tokenize,
		compile,
		select,
		outermostContext,
		sortInput,
		hasDuplicate,
	
		// Local document vars
		setDocument,
		document,
		docElem,
		documentIsHTML,
		rbuggyQSA,
		rbuggyMatches,
		matches,
		contains,
	
		// Instance-specific data
		expando = "sizzle" + 1 * new Date(),
		preferredDoc = window.document,
		dirruns = 0,
		done = 0,
		classCache = createCache(),
		tokenCache = createCache(),
		compilerCache = createCache(),
		sortOrder = function( a, b ) {
			if ( a === b ) {
				hasDuplicate = true;
			}
			return 0;
		},
	
		// Instance methods
		hasOwn = ({}).hasOwnProperty,
		arr = [],
		pop = arr.pop,
		push_native = arr.push,
		push = arr.push,
		slice = arr.slice,
		// Use a stripped-down indexOf as it's faster than native
		// https://jsperf.com/thor-indexof-vs-for/5
		indexOf = function( list, elem ) {
			var i = 0,
				len = list.length;
			for ( ; i < len; i++ ) {
				if ( list[i] === elem ) {
					return i;
				}
			}
			return -1;
		},
	
		booleans = "checked|selected|async|autofocus|autoplay|controls|defer|disabled|hidden|ismap|loop|multiple|open|readonly|required|scoped",
	
		// Regular expressions
	
		// http://www.w3.org/TR/css3-selectors/#whitespace
		whitespace = "[\\x20\\t\\r\\n\\f]",
	
		// http://www.w3.org/TR/CSS21/syndata.html#value-def-identifier
		identifier = "(?:\\\\.|[\\w-]|[^\0-\\xa0])+",
	
		// Attribute selectors: http://www.w3.org/TR/selectors/#attribute-selectors
		attributes = "\\[" + whitespace + "*(" + identifier + ")(?:" + whitespace +
			// Operator (capture 2)
			"*([*^$|!~]?=)" + whitespace +
			// "Attribute values must be CSS identifiers [capture 5] or strings [capture 3 or capture 4]"
			"*(?:'((?:\\\\.|[^\\\\'])*)'|\"((?:\\\\.|[^\\\\\"])*)\"|(" + identifier + "))|)" + whitespace +
			"*\\]",
	
		pseudos = ":(" + identifier + ")(?:\\((" +
			// To reduce the number of selectors needing tokenize in the preFilter, prefer arguments:
			// 1. quoted (capture 3; capture 4 or capture 5)
			"('((?:\\\\.|[^\\\\'])*)'|\"((?:\\\\.|[^\\\\\"])*)\")|" +
			// 2. simple (capture 6)
			"((?:\\\\.|[^\\\\()[\\]]|" + attributes + ")*)|" +
			// 3. anything else (capture 2)
			".*" +
			")\\)|)",
	
		// Leading and non-escaped trailing whitespace, capturing some non-whitespace characters preceding the latter
		rwhitespace = new RegExp( whitespace + "+", "g" ),
		rtrim = new RegExp( "^" + whitespace + "+|((?:^|[^\\\\])(?:\\\\.)*)" + whitespace + "+$", "g" ),
	
		rcomma = new RegExp( "^" + whitespace + "*," + whitespace + "*" ),
		rcombinators = new RegExp( "^" + whitespace + "*([>+~]|" + whitespace + ")" + whitespace + "*" ),
	
		rattributeQuotes = new RegExp( "=" + whitespace + "*([^\\]'\"]*?)" + whitespace + "*\\]", "g" ),
	
		rpseudo = new RegExp( pseudos ),
		ridentifier = new RegExp( "^" + identifier + "$" ),
	
		matchExpr = {
			"ID": new RegExp( "^#(" + identifier + ")" ),
			"CLASS": new RegExp( "^\\.(" + identifier + ")" ),
			"TAG": new RegExp( "^(" + identifier + "|[*])" ),
			"ATTR": new RegExp( "^" + attributes ),
			"PSEUDO": new RegExp( "^" + pseudos ),
			"CHILD": new RegExp( "^:(only|first|last|nth|nth-last)-(child|of-type)(?:\\(" + whitespace +
				"*(even|odd|(([+-]|)(\\d*)n|)" + whitespace + "*(?:([+-]|)" + whitespace +
				"*(\\d+)|))" + whitespace + "*\\)|)", "i" ),
			"bool": new RegExp( "^(?:" + booleans + ")$", "i" ),
			// For use in libraries implementing .is()
			// We use this for POS matching in `select`
			"needsContext": new RegExp( "^" + whitespace + "*[>+~]|:(even|odd|eq|gt|lt|nth|first|last)(?:\\(" +
				whitespace + "*((?:-\\d)?\\d*)" + whitespace + "*\\)|)(?=[^-]|$)", "i" )
		},
	
		rinputs = /^(?:input|select|textarea|button)$/i,
		rheader = /^h\d$/i,
	
		rnative = /^[^{]+\{\s*\[native \w/,
	
		// Easily-parseable/retrievable ID or TAG or CLASS selectors
		rquickExpr = /^(?:#([\w-]+)|(\w+)|\.([\w-]+))$/,
	
		rsibling = /[+~]/,
	
		// CSS escapes
		// http://www.w3.org/TR/CSS21/syndata.html#escaped-characters
		runescape = new RegExp( "\\\\([\\da-f]{1,6}" + whitespace + "?|(" + whitespace + ")|.)", "ig" ),
		funescape = function( _, escaped, escapedWhitespace ) {
			var high = "0x" + escaped - 0x10000;
			// NaN means non-codepoint
			// Support: Firefox<24
			// Workaround erroneous numeric interpretation of +"0x"
			return high !== high || escapedWhitespace ?
				escaped :
				high < 0 ?
					// BMP codepoint
					String.fromCharCode( high + 0x10000 ) :
					// Supplemental Plane codepoint (surrogate pair)
					String.fromCharCode( high >> 10 | 0xD800, high & 0x3FF | 0xDC00 );
		},
	
		// CSS string/identifier serialization
		// https://drafts.csswg.org/cssom/#common-serializing-idioms
		rcssescape = /([\0-\x1f\x7f]|^-?\d)|^-$|[^\0-\x1f\x7f-\uFFFF\w-]/g,
		fcssescape = function( ch, asCodePoint ) {
			if ( asCodePoint ) {
	
				// U+0000 NULL becomes U+FFFD REPLACEMENT CHARACTER
				if ( ch === "\0" ) {
					return "\uFFFD";
				}
	
				// Control characters and (dependent upon position) numbers get escaped as code points
				return ch.slice( 0, -1 ) + "\\" + ch.charCodeAt( ch.length - 1 ).toString( 16 ) + " ";
			}
	
			// Other potentially-special ASCII characters get backslash-escaped
			return "\\" + ch;
		},
	
		// Used for iframes
		// See setDocument()
		// Removing the function wrapper causes a "Permission Denied"
		// error in IE
		unloadHandler = function() {
			setDocument();
		},
	
		disabledAncestor = addCombinator(
			function( elem ) {
				return elem.disabled === true && ("form" in elem || "label" in elem);
			},
			{ dir: "parentNode", next: "legend" }
		);
	
	// Optimize for push.apply( _, NodeList )
	try {
		push.apply(
			(arr = slice.call( preferredDoc.childNodes )),
			preferredDoc.childNodes
		);
		// Support: Android<4.0
		// Detect silently failing push.apply
		arr[ preferredDoc.childNodes.length ].nodeType;
	} catch ( e ) {
		push = { apply: arr.length ?
	
			// Leverage slice if possible
			function( target, els ) {
				push_native.apply( target, slice.call(els) );
			} :
	
			// Support: IE<9
			// Otherwise append directly
			function( target, els ) {
				var j = target.length,
					i = 0;
				// Can't trust NodeList.length
				while ( (target[j++] = els[i++]) ) {}
				target.length = j - 1;
			}
		};
	}
	
	function Sizzle( selector, context, results, seed ) {
		var m, i, elem, nid, match, groups, newSelector,
			newContext = context && context.ownerDocument,
	
			// nodeType defaults to 9, since context defaults to document
			nodeType = context ? context.nodeType : 9;
	
		results = results || [];
	
		// Return early from calls with invalid selector or context
		if ( typeof selector !== "string" || !selector ||
			nodeType !== 1 && nodeType !== 9 && nodeType !== 11 ) {
	
			return results;
		}
	
		// Try to shortcut find operations (as opposed to filters) in HTML documents
		if ( !seed ) {
	
			if ( ( context ? context.ownerDocument || context : preferredDoc ) !== document ) {
				setDocument( context );
			}
			context = context || document;
	
			if ( documentIsHTML ) {
	
				// If the selector is sufficiently simple, try using a "get*By*" DOM method
				// (excepting DocumentFragment context, where the methods don't exist)
				if ( nodeType !== 11 && (match = rquickExpr.exec( selector )) ) {
	
					// ID selector
					if ( (m = match[1]) ) {
	
						// Document context
						if ( nodeType === 9 ) {
							if ( (elem = context.getElementById( m )) ) {
	
								// Support: IE, Opera, Webkit
								// TODO: identify versions
								// getElementById can match elements by name instead of ID
								if ( elem.id === m ) {
									results.push( elem );
									return results;
								}
							} else {
								return results;
							}
	
						// Element context
						} else {
	
							// Support: IE, Opera, Webkit
							// TODO: identify versions
							// getElementById can match elements by name instead of ID
							if ( newContext && (elem = newContext.getElementById( m )) &&
								contains( context, elem ) &&
								elem.id === m ) {
	
								results.push( elem );
								return results;
							}
						}
	
					// Type selector
					} else if ( match[2] ) {
						push.apply( results, context.getElementsByTagName( selector ) );
						return results;
	
					// Class selector
					} else if ( (m = match[3]) && support.getElementsByClassName &&
						context.getElementsByClassName ) {
	
						push.apply( results, context.getElementsByClassName( m ) );
						return results;
					}
				}
	
				// Take advantage of querySelectorAll
				if ( support.qsa &&
					!compilerCache[ selector + " " ] &&
					(!rbuggyQSA || !rbuggyQSA.test( selector )) ) {
	
					if ( nodeType !== 1 ) {
						newContext = context;
						newSelector = selector;
	
					// qSA looks outside Element context, which is not what we want
					// Thanks to Andrew Dupont for this workaround technique
					// Support: IE <=8
					// Exclude object elements
					} else if ( context.nodeName.toLowerCase() !== "object" ) {
	
						// Capture the context ID, setting it first if necessary
						if ( (nid = context.getAttribute( "id" )) ) {
							nid = nid.replace( rcssescape, fcssescape );
						} else {
							context.setAttribute( "id", (nid = expando) );
						}
	
						// Prefix every selector in the list
						groups = tokenize( selector );
						i = groups.length;
						while ( i-- ) {
							groups[i] = "#" + nid + " " + toSelector( groups[i] );
						}
						newSelector = groups.join( "," );
	
						// Expand context for sibling selectors
						newContext = rsibling.test( selector ) && testContext( context.parentNode ) ||
							context;
					}
	
					if ( newSelector ) {
						try {
							push.apply( results,
								newContext.querySelectorAll( newSelector )
							);
							return results;
						} catch ( qsaError ) {
						} finally {
							if ( nid === expando ) {
								context.removeAttribute( "id" );
							}
						}
					}
				}
			}
		}
	
		// All others
		return select( selector.replace( rtrim, "$1" ), context, results, seed );
	}
	
	/**
	 * Create key-value caches of limited size
	 * @returns {function(string, object)} Returns the Object data after storing it on itself with
	 *	property name the (space-suffixed) string and (if the cache is larger than Expr.cacheLength)
	 *	deleting the oldest entry
	 */
	function createCache() {
		var keys = [];
	
		function cache( key, value ) {
			// Use (key + " ") to avoid collision with native prototype properties (see Issue #157)
			if ( keys.push( key + " " ) > Expr.cacheLength ) {
				// Only keep the most recent entries
				delete cache[ keys.shift() ];
			}
			return (cache[ key + " " ] = value);
		}
		return cache;
	}
	
	/**
	 * Mark a function for special use by Sizzle
	 * @param {Function} fn The function to mark
	 */
	function markFunction( fn ) {
		fn[ expando ] = true;
		return fn;
	}
	
	/**
	 * Support testing using an element
	 * @param {Function} fn Passed the created element and returns a boolean result
	 */
	function assert( fn ) {
		var el = document.createElement("fieldset");
	
		try {
			return !!fn( el );
		} catch (e) {
			return false;
		} finally {
			// Remove from its parent by default
			if ( el.parentNode ) {
				el.parentNode.removeChild( el );
			}
			// release memory in IE
			el = null;
		}
	}
	
	/**
	 * Adds the same handler for all of the specified attrs
	 * @param {String} attrs Pipe-separated list of attributes
	 * @param {Function} handler The method that will be applied
	 */
	function addHandle( attrs, handler ) {
		var arr = attrs.split("|"),
			i = arr.length;
	
		while ( i-- ) {
			Expr.attrHandle[ arr[i] ] = handler;
		}
	}
	
	/**
	 * Checks document order of two siblings
	 * @param {Element} a
	 * @param {Element} b
	 * @returns {Number} Returns less than 0 if a precedes b, greater than 0 if a follows b
	 */
	function siblingCheck( a, b ) {
		var cur = b && a,
			diff = cur && a.nodeType === 1 && b.nodeType === 1 &&
				a.sourceIndex - b.sourceIndex;
	
		// Use IE sourceIndex if available on both nodes
		if ( diff ) {
			return diff;
		}
	
		// Check if b follows a
		if ( cur ) {
			while ( (cur = cur.nextSibling) ) {
				if ( cur === b ) {
					return -1;
				}
			}
		}
	
		return a ? 1 : -1;
	}
	
	/**
	 * Returns a function to use in pseudos for input types
	 * @param {String} type
	 */
	function createInputPseudo( type ) {
		return function( elem ) {
			var name = elem.nodeName.toLowerCase();
			return name === "input" && elem.type === type;
		};
	}
	
	/**
	 * Returns a function to use in pseudos for buttons
	 * @param {String} type
	 */
	function createButtonPseudo( type ) {
		return function( elem ) {
			var name = elem.nodeName.toLowerCase();
			return (name === "input" || name === "button") && elem.type === type;
		};
	}
	
	/**
	 * Returns a function to use in pseudos for :enabled/:disabled
	 * @param {Boolean} disabled true for :disabled; false for :enabled
	 */
	function createDisabledPseudo( disabled ) {
	
		// Known :disabled false positives: fieldset[disabled] > legend:nth-of-type(n+2) :can-disable
		return function( elem ) {
	
			// Only certain elements can match :enabled or :disabled
			// https://html.spec.whatwg.org/multipage/scripting.html#selector-enabled
			// https://html.spec.whatwg.org/multipage/scripting.html#selector-disabled
			if ( "form" in elem ) {
	
				// Check for inherited disabledness on relevant non-disabled elements:
				// * listed form-associated elements in a disabled fieldset
				//   https://html.spec.whatwg.org/multipage/forms.html#category-listed
				//   https://html.spec.whatwg.org/multipage/forms.html#concept-fe-disabled
				// * option elements in a disabled optgroup
				//   https://html.spec.whatwg.org/multipage/forms.html#concept-option-disabled
				// All such elements have a "form" property.
				if ( elem.parentNode && elem.disabled === false ) {
	
					// Option elements defer to a parent optgroup if present
					if ( "label" in elem ) {
						if ( "label" in elem.parentNode ) {
							return elem.parentNode.disabled === disabled;
						} else {
							return elem.disabled === disabled;
						}
					}
	
					// Support: IE 6 - 11
					// Use the isDisabled shortcut property to check for disabled fieldset ancestors
					return elem.isDisabled === disabled ||
	
						// Where there is no isDisabled, check manually
						/* jshint -W018 */
						elem.isDisabled !== !disabled &&
							disabledAncestor( elem ) === disabled;
				}
	
				return elem.disabled === disabled;
	
			// Try to winnow out elements that can't be disabled before trusting the disabled property.
			// Some victims get caught in our net (label, legend, menu, track), but it shouldn't
			// even exist on them, let alone have a boolean value.
			} else if ( "label" in elem ) {
				return elem.disabled === disabled;
			}
	
			// Remaining elements are neither :enabled nor :disabled
			return false;
		};
	}
	
	/**
	 * Returns a function to use in pseudos for positionals
	 * @param {Function} fn
	 */
	function createPositionalPseudo( fn ) {
		return markFunction(function( argument ) {
			argument = +argument;
			return markFunction(function( seed, matches ) {
				var j,
					matchIndexes = fn( [], seed.length, argument ),
					i = matchIndexes.length;
	
				// Match elements found at the specified indexes
				while ( i-- ) {
					if ( seed[ (j = matchIndexes[i]) ] ) {
						seed[j] = !(matches[j] = seed[j]);
					}
				}
			});
		});
	}
	
	/**
	 * Checks a node for validity as a Sizzle context
	 * @param {Element|Object=} context
	 * @returns {Element|Object|Boolean} The input node if acceptable, otherwise a falsy value
	 */
	function testContext( context ) {
		return context && typeof context.getElementsByTagName !== "undefined" && context;
	}
	
	// Expose support vars for convenience
	support = Sizzle.support = {};
	
	/**
	 * Detects XML nodes
	 * @param {Element|Object} elem An element or a document
	 * @returns {Boolean} True iff elem is a non-HTML XML node
	 */
	isXML = Sizzle.isXML = function( elem ) {
		// documentElement is verified for cases where it doesn't yet exist
		// (such as loading iframes in IE - #4833)
		var documentElement = elem && (elem.ownerDocument || elem).documentElement;
		return documentElement ? documentElement.nodeName !== "HTML" : false;
	};
	
	/**
	 * Sets document-related variables once based on the current document
	 * @param {Element|Object} [doc] An element or document object to use to set the document
	 * @returns {Object} Returns the current document
	 */
	setDocument = Sizzle.setDocument = function( node ) {
		var hasCompare, subWindow,
			doc = node ? node.ownerDocument || node : preferredDoc;
	
		// Return early if doc is invalid or already selected
		if ( doc === document || doc.nodeType !== 9 || !doc.documentElement ) {
			return document;
		}
	
		// Update global variables
		document = doc;
		docElem = document.documentElement;
		documentIsHTML = !isXML( document );
	
		// Support: IE 9-11, Edge
		// Accessing iframe documents after unload throws "permission denied" errors (jQuery #13936)
		if ( preferredDoc !== document &&
			(subWindow = document.defaultView) && subWindow.top !== subWindow ) {
	
			// Support: IE 11, Edge
			if ( subWindow.addEventListener ) {
				subWindow.addEventListener( "unload", unloadHandler, false );
	
			// Support: IE 9 - 10 only
			} else if ( subWindow.attachEvent ) {
				subWindow.attachEvent( "onunload", unloadHandler );
			}
		}
	
		/* Attributes
		---------------------------------------------------------------------- */
	
		// Support: IE<8
		// Verify that getAttribute really returns attributes and not properties
		// (excepting IE8 booleans)
		support.attributes = assert(function( el ) {
			el.className = "i";
			return !el.getAttribute("className");
		});
	
		/* getElement(s)By*
		---------------------------------------------------------------------- */
	
		// Check if getElementsByTagName("*") returns only elements
		support.getElementsByTagName = assert(function( el ) {
			el.appendChild( document.createComment("") );
			return !el.getElementsByTagName("*").length;
		});
	
		// Support: IE<9
		support.getElementsByClassName = rnative.test( document.getElementsByClassName );
	
		// Support: IE<10
		// Check if getElementById returns elements by name
		// The broken getElementById methods don't pick up programmatically-set names,
		// so use a roundabout getElementsByName test
		support.getById = assert(function( el ) {
			docElem.appendChild( el ).id = expando;
			return !document.getElementsByName || !document.getElementsByName( expando ).length;
		});
	
		// ID filter and find
		if ( support.getById ) {
			Expr.filter["ID"] = function( id ) {
				var attrId = id.replace( runescape, funescape );
				return function( elem ) {
					return elem.getAttribute("id") === attrId;
				};
			};
			Expr.find["ID"] = function( id, context ) {
				if ( typeof context.getElementById !== "undefined" && documentIsHTML ) {
					var elem = context.getElementById( id );
					return elem ? [ elem ] : [];
				}
			};
		} else {
			Expr.filter["ID"] =  function( id ) {
				var attrId = id.replace( runescape, funescape );
				return function( elem ) {
					var node = typeof elem.getAttributeNode !== "undefined" &&
						elem.getAttributeNode("id");
					return node && node.value === attrId;
				};
			};
	
			// Support: IE 6 - 7 only
			// getElementById is not reliable as a find shortcut
			Expr.find["ID"] = function( id, context ) {
				if ( typeof context.getElementById !== "undefined" && documentIsHTML ) {
					var node, i, elems,
						elem = context.getElementById( id );
	
					if ( elem ) {
	
						// Verify the id attribute
						node = elem.getAttributeNode("id");
						if ( node && node.value === id ) {
							return [ elem ];
						}
	
						// Fall back on getElementsByName
						elems = context.getElementsByName( id );
						i = 0;
						while ( (elem = elems[i++]) ) {
							node = elem.getAttributeNode("id");
							if ( node && node.value === id ) {
								return [ elem ];
							}
						}
					}
	
					return [];
				}
			};
		}
	
		// Tag
		Expr.find["TAG"] = support.getElementsByTagName ?
			function( tag, context ) {
				if ( typeof context.getElementsByTagName !== "undefined" ) {
					return context.getElementsByTagName( tag );
	
				// DocumentFragment nodes don't have gEBTN
				} else if ( support.qsa ) {
					return context.querySelectorAll( tag );
				}
			} :
	
			function( tag, context ) {
				var elem,
					tmp = [],
					i = 0,
					// By happy coincidence, a (broken) gEBTN appears on DocumentFragment nodes too
					results = context.getElementsByTagName( tag );
	
				// Filter out possible comments
				if ( tag === "*" ) {
					while ( (elem = results[i++]) ) {
						if ( elem.nodeType === 1 ) {
							tmp.push( elem );
						}
					}
	
					return tmp;
				}
				return results;
			};
	
		// Class
		Expr.find["CLASS"] = support.getElementsByClassName && function( className, context ) {
			if ( typeof context.getElementsByClassName !== "undefined" && documentIsHTML ) {
				return context.getElementsByClassName( className );
			}
		};
	
		/* QSA/matchesSelector
		---------------------------------------------------------------------- */
	
		// QSA and matchesSelector support
	
		// matchesSelector(:active) reports false when true (IE9/Opera 11.5)
		rbuggyMatches = [];
	
		// qSa(:focus) reports false when true (Chrome 21)
		// We allow this because of a bug in IE8/9 that throws an error
		// whenever `document.activeElement` is accessed on an iframe
		// So, we allow :focus to pass through QSA all the time to avoid the IE error
		// See https://bugs.jquery.com/ticket/13378
		rbuggyQSA = [];
	
		if ( (support.qsa = rnative.test( document.querySelectorAll )) ) {
			// Build QSA regex
			// Regex strategy adopted from Diego Perini
			assert(function( el ) {
				// Select is set to empty string on purpose
				// This is to test IE's treatment of not explicitly
				// setting a boolean content attribute,
				// since its presence should be enough
				// https://bugs.jquery.com/ticket/12359
				docElem.appendChild( el ).innerHTML = "<a id='" + expando + "'></a>" +
					"<select id='" + expando + "-\r\\' msallowcapture=''>" +
					"<option selected=''></option></select>";
	
				// Support: IE8, Opera 11-12.16
				// Nothing should be selected when empty strings follow ^= or $= or *=
				// The test attribute must be unknown in Opera but "safe" for WinRT
				// https://msdn.microsoft.com/en-us/library/ie/hh465388.aspx#attribute_section
				if ( el.querySelectorAll("[msallowcapture^='']").length ) {
					rbuggyQSA.push( "[*^$]=" + whitespace + "*(?:''|\"\")" );
				}
	
				// Support: IE8
				// Boolean attributes and "value" are not treated correctly
				if ( !el.querySelectorAll("[selected]").length ) {
					rbuggyQSA.push( "\\[" + whitespace + "*(?:value|" + booleans + ")" );
				}
	
				// Support: Chrome<29, Android<4.4, Safari<7.0+, iOS<7.0+, PhantomJS<1.9.8+
				if ( !el.querySelectorAll( "[id~=" + expando + "-]" ).length ) {
					rbuggyQSA.push("~=");
				}
	
				// Webkit/Opera - :checked should return selected option elements
				// http://www.w3.org/TR/2011/REC-css3-selectors-20110929/#checked
				// IE8 throws error here and will not see later tests
				if ( !el.querySelectorAll(":checked").length ) {
					rbuggyQSA.push(":checked");
				}
	
				// Support: Safari 8+, iOS 8+
				// https://bugs.webkit.org/show_bug.cgi?id=136851
				// In-page `selector#id sibling-combinator selector` fails
				if ( !el.querySelectorAll( "a#" + expando + "+*" ).length ) {
					rbuggyQSA.push(".#.+[+~]");
				}
			});
	
			assert(function( el ) {
				el.innerHTML = "<a href='' disabled='disabled'></a>" +
					"<select disabled='disabled'><option/></select>";
	
				// Support: Windows 8 Native Apps
				// The type and name attributes are restricted during .innerHTML assignment
				var input = document.createElement("input");
				input.setAttribute( "type", "hidden" );
				el.appendChild( input ).setAttribute( "name", "D" );
	
				// Support: IE8
				// Enforce case-sensitivity of name attribute
				if ( el.querySelectorAll("[name=d]").length ) {
					rbuggyQSA.push( "name" + whitespace + "*[*^$|!~]?=" );
				}
	
				// FF 3.5 - :enabled/:disabled and hidden elements (hidden elements are still enabled)
				// IE8 throws error here and will not see later tests
				if ( el.querySelectorAll(":enabled").length !== 2 ) {
					rbuggyQSA.push( ":enabled", ":disabled" );
				}
	
				// Support: IE9-11+
				// IE's :disabled selector does not pick up the children of disabled fieldsets
				docElem.appendChild( el ).disabled = true;
				if ( el.querySelectorAll(":disabled").length !== 2 ) {
					rbuggyQSA.push( ":enabled", ":disabled" );
				}
	
				// Opera 10-11 does not throw on post-comma invalid pseudos
				el.querySelectorAll("*,:x");
				rbuggyQSA.push(",.*:");
			});
		}
	
		if ( (support.matchesSelector = rnative.test( (matches = docElem.matches ||
			docElem.webkitMatchesSelector ||
			docElem.mozMatchesSelector ||
			docElem.oMatchesSelector ||
			docElem.msMatchesSelector) )) ) {
	
			assert(function( el ) {
				// Check to see if it's possible to do matchesSelector
				// on a disconnected node (IE 9)
				support.disconnectedMatch = matches.call( el, "*" );
	
				// This should fail with an exception
				// Gecko does not error, returns false instead
				matches.call( el, "[s!='']:x" );
				rbuggyMatches.push( "!=", pseudos );
			});
		}
	
		rbuggyQSA = rbuggyQSA.length && new RegExp( rbuggyQSA.join("|") );
		rbuggyMatches = rbuggyMatches.length && new RegExp( rbuggyMatches.join("|") );
	
		/* Contains
		---------------------------------------------------------------------- */
		hasCompare = rnative.test( docElem.compareDocumentPosition );
	
		// Element contains another
		// Purposefully self-exclusive
		// As in, an element does not contain itself
		contains = hasCompare || rnative.test( docElem.contains ) ?
			function( a, b ) {
				var adown = a.nodeType === 9 ? a.documentElement : a,
					bup = b && b.parentNode;
				return a === bup || !!( bup && bup.nodeType === 1 && (
					adown.contains ?
						adown.contains( bup ) :
						a.compareDocumentPosition && a.compareDocumentPosition( bup ) & 16
				));
			} :
			function( a, b ) {
				if ( b ) {
					while ( (b = b.parentNode) ) {
						if ( b === a ) {
							return true;
						}
					}
				}
				return false;
			};
	
		/* Sorting
		---------------------------------------------------------------------- */
	
		// Document order sorting
		sortOrder = hasCompare ?
		function( a, b ) {
	
			// Flag for duplicate removal
			if ( a === b ) {
				hasDuplicate = true;
				return 0;
			}
	
			// Sort on method existence if only one input has compareDocumentPosition
			var compare = !a.compareDocumentPosition - !b.compareDocumentPosition;
			if ( compare ) {
				return compare;
			}
	
			// Calculate position if both inputs belong to the same document
			compare = ( a.ownerDocument || a ) === ( b.ownerDocument || b ) ?
				a.compareDocumentPosition( b ) :
	
				// Otherwise we know they are disconnected
				1;
	
			// Disconnected nodes
			if ( compare & 1 ||
				(!support.sortDetached && b.compareDocumentPosition( a ) === compare) ) {
	
				// Choose the first element that is related to our preferred document
				if ( a === document || a.ownerDocument === preferredDoc && contains(preferredDoc, a) ) {
					return -1;
				}
				if ( b === document || b.ownerDocument === preferredDoc && contains(preferredDoc, b) ) {
					return 1;
				}
	
				// Maintain original order
				return sortInput ?
					( indexOf( sortInput, a ) - indexOf( sortInput, b ) ) :
					0;
			}
	
			return compare & 4 ? -1 : 1;
		} :
		function( a, b ) {
			// Exit early if the nodes are identical
			if ( a === b ) {
				hasDuplicate = true;
				return 0;
			}
	
			var cur,
				i = 0,
				aup = a.parentNode,
				bup = b.parentNode,
				ap = [ a ],
				bp = [ b ];
	
			// Parentless nodes are either documents or disconnected
			if ( !aup || !bup ) {
				return a === document ? -1 :
					b === document ? 1 :
					aup ? -1 :
					bup ? 1 :
					sortInput ?
					( indexOf( sortInput, a ) - indexOf( sortInput, b ) ) :
					0;
	
			// If the nodes are siblings, we can do a quick check
			} else if ( aup === bup ) {
				return siblingCheck( a, b );
			}
	
			// Otherwise we need full lists of their ancestors for comparison
			cur = a;
			while ( (cur = cur.parentNode) ) {
				ap.unshift( cur );
			}
			cur = b;
			while ( (cur = cur.parentNode) ) {
				bp.unshift( cur );
			}
	
			// Walk down the tree looking for a discrepancy
			while ( ap[i] === bp[i] ) {
				i++;
			}
	
			return i ?
				// Do a sibling check if the nodes have a common ancestor
				siblingCheck( ap[i], bp[i] ) :
	
				// Otherwise nodes in our document sort first
				ap[i] === preferredDoc ? -1 :
				bp[i] === preferredDoc ? 1 :
				0;
		};
	
		return document;
	};
	
	Sizzle.matches = function( expr, elements ) {
		return Sizzle( expr, null, null, elements );
	};
	
	Sizzle.matchesSelector = function( elem, expr ) {
		// Set document vars if needed
		if ( ( elem.ownerDocument || elem ) !== document ) {
			setDocument( elem );
		}
	
		// Make sure that attribute selectors are quoted
		expr = expr.replace( rattributeQuotes, "='$1']" );
	
		if ( support.matchesSelector && documentIsHTML &&
			!compilerCache[ expr + " " ] &&
			( !rbuggyMatches || !rbuggyMatches.test( expr ) ) &&
			( !rbuggyQSA     || !rbuggyQSA.test( expr ) ) ) {
	
			try {
				var ret = matches.call( elem, expr );
	
				// IE 9's matchesSelector returns false on disconnected nodes
				if ( ret || support.disconnectedMatch ||
						// As well, disconnected nodes are said to be in a document
						// fragment in IE 9
						elem.document && elem.document.nodeType !== 11 ) {
					return ret;
				}
			} catch (e) {}
		}
	
		return Sizzle( expr, document, null, [ elem ] ).length > 0;
	};
	
	Sizzle.contains = function( context, elem ) {
		// Set document vars if needed
		if ( ( context.ownerDocument || context ) !== document ) {
			setDocument( context );
		}
		return contains( context, elem );
	};
	
	Sizzle.attr = function( elem, name ) {
		// Set document vars if needed
		if ( ( elem.ownerDocument || elem ) !== document ) {
			setDocument( elem );
		}
	
		var fn = Expr.attrHandle[ name.toLowerCase() ],
			// Don't get fooled by Object.prototype properties (jQuery #13807)
			val = fn && hasOwn.call( Expr.attrHandle, name.toLowerCase() ) ?
				fn( elem, name, !documentIsHTML ) :
				undefined;
	
		return val !== undefined ?
			val :
			support.attributes || !documentIsHTML ?
				elem.getAttribute( name ) :
				(val = elem.getAttributeNode(name)) && val.specified ?
					val.value :
					null;
	};
	
	Sizzle.escape = function( sel ) {
		return (sel + "").replace( rcssescape, fcssescape );
	};
	
	Sizzle.error = function( msg ) {
		throw new Error( "Syntax error, unrecognized expression: " + msg );
	};
	
	/**
	 * Document sorting and removing duplicates
	 * @param {ArrayLike} results
	 */
	Sizzle.uniqueSort = function( results ) {
		var elem,
			duplicates = [],
			j = 0,
			i = 0;
	
		// Unless we *know* we can detect duplicates, assume their presence
		hasDuplicate = !support.detectDuplicates;
		sortInput = !support.sortStable && results.slice( 0 );
		results.sort( sortOrder );
	
		if ( hasDuplicate ) {
			while ( (elem = results[i++]) ) {
				if ( elem === results[ i ] ) {
					j = duplicates.push( i );
				}
			}
			while ( j-- ) {
				results.splice( duplicates[ j ], 1 );
			}
		}
	
		// Clear input after sorting to release objects
		// See https://github.com/jquery/sizzle/pull/225
		sortInput = null;
	
		return results;
	};
	
	/**
	 * Utility function for retrieving the text value of an array of DOM nodes
	 * @param {Array|Element} elem
	 */
	getText = Sizzle.getText = function( elem ) {
		var node,
			ret = "",
			i = 0,
			nodeType = elem.nodeType;
	
		if ( !nodeType ) {
			// If no nodeType, this is expected to be an array
			while ( (node = elem[i++]) ) {
				// Do not traverse comment nodes
				ret += getText( node );
			}
		} else if ( nodeType === 1 || nodeType === 9 || nodeType === 11 ) {
			// Use textContent for elements
			// innerText usage removed for consistency of new lines (jQuery #11153)
			if ( typeof elem.textContent === "string" ) {
				return elem.textContent;
			} else {
				// Traverse its children
				for ( elem = elem.firstChild; elem; elem = elem.nextSibling ) {
					ret += getText( elem );
				}
			}
		} else if ( nodeType === 3 || nodeType === 4 ) {
			return elem.nodeValue;
		}
		// Do not include comment or processing instruction nodes
	
		return ret;
	};
	
	Expr = Sizzle.selectors = {
	
		// Can be adjusted by the user
		cacheLength: 50,
	
		createPseudo: markFunction,
	
		match: matchExpr,
	
		attrHandle: {},
	
		find: {},
	
		relative: {
			">": { dir: "parentNode", first: true },
			" ": { dir: "parentNode" },
			"+": { dir: "previousSibling", first: true },
			"~": { dir: "previousSibling" }
		},
	
		preFilter: {
			"ATTR": function( match ) {
				match[1] = match[1].replace( runescape, funescape );
	
				// Move the given value to match[3] whether quoted or unquoted
				match[3] = ( match[3] || match[4] || match[5] || "" ).replace( runescape, funescape );
	
				if ( match[2] === "~=" ) {
					match[3] = " " + match[3] + " ";
				}
	
				return match.slice( 0, 4 );
			},
	
			"CHILD": function( match ) {
				/* matches from matchExpr["CHILD"]
					1 type (only|nth|...)
					2 what (child|of-type)
					3 argument (even|odd|\d*|\d*n([+-]\d+)?|...)
					4 xn-component of xn+y argument ([+-]?\d*n|)
					5 sign of xn-component
					6 x of xn-component
					7 sign of y-component
					8 y of y-component
				*/
				match[1] = match[1].toLowerCase();
	
				if ( match[1].slice( 0, 3 ) === "nth" ) {
					// nth-* requires argument
					if ( !match[3] ) {
						Sizzle.error( match[0] );
					}
	
					// numeric x and y parameters for Expr.filter.CHILD
					// remember that false/true cast respectively to 0/1
					match[4] = +( match[4] ? match[5] + (match[6] || 1) : 2 * ( match[3] === "even" || match[3] === "odd" ) );
					match[5] = +( ( match[7] + match[8] ) || match[3] === "odd" );
	
				// other types prohibit arguments
				} else if ( match[3] ) {
					Sizzle.error( match[0] );
				}
	
				return match;
			},
	
			"PSEUDO": function( match ) {
				var excess,
					unquoted = !match[6] && match[2];
	
				if ( matchExpr["CHILD"].test( match[0] ) ) {
					return null;
				}
	
				// Accept quoted arguments as-is
				if ( match[3] ) {
					match[2] = match[4] || match[5] || "";
	
				// Strip excess characters from unquoted arguments
				} else if ( unquoted && rpseudo.test( unquoted ) &&
					// Get excess from tokenize (recursively)
					(excess = tokenize( unquoted, true )) &&
					// advance to the next closing parenthesis
					(excess = unquoted.indexOf( ")", unquoted.length - excess ) - unquoted.length) ) {
	
					// excess is a negative index
					match[0] = match[0].slice( 0, excess );
					match[2] = unquoted.slice( 0, excess );
				}
	
				// Return only captures needed by the pseudo filter method (type and argument)
				return match.slice( 0, 3 );
			}
		},
	
		filter: {
	
			"TAG": function( nodeNameSelector ) {
				var nodeName = nodeNameSelector.replace( runescape, funescape ).toLowerCase();
				return nodeNameSelector === "*" ?
					function() { return true; } :
					function( elem ) {
						return elem.nodeName && elem.nodeName.toLowerCase() === nodeName;
					};
			},
	
			"CLASS": function( className ) {
				var pattern = classCache[ className + " " ];
	
				return pattern ||
					(pattern = new RegExp( "(^|" + whitespace + ")" + className + "(" + whitespace + "|$)" )) &&
					classCache( className, function( elem ) {
						return pattern.test( typeof elem.className === "string" && elem.className || typeof elem.getAttribute !== "undefined" && elem.getAttribute("class") || "" );
					});
			},
	
			"ATTR": function( name, operator, check ) {
				return function( elem ) {
					var result = Sizzle.attr( elem, name );
	
					if ( result == null ) {
						return operator === "!=";
					}
					if ( !operator ) {
						return true;
					}
	
					result += "";
	
					return operator === "=" ? result === check :
						operator === "!=" ? result !== check :
						operator === "^=" ? check && result.indexOf( check ) === 0 :
						operator === "*=" ? check && result.indexOf( check ) > -1 :
						operator === "$=" ? check && result.slice( -check.length ) === check :
						operator === "~=" ? ( " " + result.replace( rwhitespace, " " ) + " " ).indexOf( check ) > -1 :
						operator === "|=" ? result === check || result.slice( 0, check.length + 1 ) === check + "-" :
						false;
				};
			},
	
			"CHILD": function( type, what, argument, first, last ) {
				var simple = type.slice( 0, 3 ) !== "nth",
					forward = type.slice( -4 ) !== "last",
					ofType = what === "of-type";
	
				return first === 1 && last === 0 ?
	
					// Shortcut for :nth-*(n)
					function( elem ) {
						return !!elem.parentNode;
					} :
	
					function( elem, context, xml ) {
						var cache, uniqueCache, outerCache, node, nodeIndex, start,
							dir = simple !== forward ? "nextSibling" : "previousSibling",
							parent = elem.parentNode,
							name = ofType && elem.nodeName.toLowerCase(),
							useCache = !xml && !ofType,
							diff = false;
	
						if ( parent ) {
	
							// :(first|last|only)-(child|of-type)
							if ( simple ) {
								while ( dir ) {
									node = elem;
									while ( (node = node[ dir ]) ) {
										if ( ofType ?
											node.nodeName.toLowerCase() === name :
											node.nodeType === 1 ) {
	
											return false;
										}
									}
									// Reverse direction for :only-* (if we haven't yet done so)
									start = dir = type === "only" && !start && "nextSibling";
								}
								return true;
							}
	
							start = [ forward ? parent.firstChild : parent.lastChild ];
	
							// non-xml :nth-child(...) stores cache data on `parent`
							if ( forward && useCache ) {
	
								// Seek `elem` from a previously-cached index
	
								// ...in a gzip-friendly way
								node = parent;
								outerCache = node[ expando ] || (node[ expando ] = {});
	
								// Support: IE <9 only
								// Defend against cloned attroperties (jQuery gh-1709)
								uniqueCache = outerCache[ node.uniqueID ] ||
									(outerCache[ node.uniqueID ] = {});
	
								cache = uniqueCache[ type ] || [];
								nodeIndex = cache[ 0 ] === dirruns && cache[ 1 ];
								diff = nodeIndex && cache[ 2 ];
								node = nodeIndex && parent.childNodes[ nodeIndex ];
	
								while ( (node = ++nodeIndex && node && node[ dir ] ||
	
									// Fallback to seeking `elem` from the start
									(diff = nodeIndex = 0) || start.pop()) ) {
	
									// When found, cache indexes on `parent` and break
									if ( node.nodeType === 1 && ++diff && node === elem ) {
										uniqueCache[ type ] = [ dirruns, nodeIndex, diff ];
										break;
									}
								}
	
							} else {
								// Use previously-cached element index if available
								if ( useCache ) {
									// ...in a gzip-friendly way
									node = elem;
									outerCache = node[ expando ] || (node[ expando ] = {});
	
									// Support: IE <9 only
									// Defend against cloned attroperties (jQuery gh-1709)
									uniqueCache = outerCache[ node.uniqueID ] ||
										(outerCache[ node.uniqueID ] = {});
	
									cache = uniqueCache[ type ] || [];
									nodeIndex = cache[ 0 ] === dirruns && cache[ 1 ];
									diff = nodeIndex;
								}
	
								// xml :nth-child(...)
								// or :nth-last-child(...) or :nth(-last)?-of-type(...)
								if ( diff === false ) {
									// Use the same loop as above to seek `elem` from the start
									while ( (node = ++nodeIndex && node && node[ dir ] ||
										(diff = nodeIndex = 0) || start.pop()) ) {
	
										if ( ( ofType ?
											node.nodeName.toLowerCase() === name :
											node.nodeType === 1 ) &&
											++diff ) {
	
											// Cache the index of each encountered element
											if ( useCache ) {
												outerCache = node[ expando ] || (node[ expando ] = {});
	
												// Support: IE <9 only
												// Defend against cloned attroperties (jQuery gh-1709)
												uniqueCache = outerCache[ node.uniqueID ] ||
													(outerCache[ node.uniqueID ] = {});
	
												uniqueCache[ type ] = [ dirruns, diff ];
											}
	
											if ( node === elem ) {
												break;
											}
										}
									}
								}
							}
	
							// Incorporate the offset, then check against cycle size
							diff -= last;
							return diff === first || ( diff % first === 0 && diff / first >= 0 );
						}
					};
			},
	
			"PSEUDO": function( pseudo, argument ) {
				// pseudo-class names are case-insensitive
				// http://www.w3.org/TR/selectors/#pseudo-classes
				// Prioritize by case sensitivity in case custom pseudos are added with uppercase letters
				// Remember that setFilters inherits from pseudos
				var args,
					fn = Expr.pseudos[ pseudo ] || Expr.setFilters[ pseudo.toLowerCase() ] ||
						Sizzle.error( "unsupported pseudo: " + pseudo );
	
				// The user may use createPseudo to indicate that
				// arguments are needed to create the filter function
				// just as Sizzle does
				if ( fn[ expando ] ) {
					return fn( argument );
				}
	
				// But maintain support for old signatures
				if ( fn.length > 1 ) {
					args = [ pseudo, pseudo, "", argument ];
					return Expr.setFilters.hasOwnProperty( pseudo.toLowerCase() ) ?
						markFunction(function( seed, matches ) {
							var idx,
								matched = fn( seed, argument ),
								i = matched.length;
							while ( i-- ) {
								idx = indexOf( seed, matched[i] );
								seed[ idx ] = !( matches[ idx ] = matched[i] );
							}
						}) :
						function( elem ) {
							return fn( elem, 0, args );
						};
				}
	
				return fn;
			}
		},
	
		pseudos: {
			// Potentially complex pseudos
			"not": markFunction(function( selector ) {
				// Trim the selector passed to compile
				// to avoid treating leading and trailing
				// spaces as combinators
				var input = [],
					results = [],
					matcher = compile( selector.replace( rtrim, "$1" ) );
	
				return matcher[ expando ] ?
					markFunction(function( seed, matches, context, xml ) {
						var elem,
							unmatched = matcher( seed, null, xml, [] ),
							i = seed.length;
	
						// Match elements unmatched by `matcher`
						while ( i-- ) {
							if ( (elem = unmatched[i]) ) {
								seed[i] = !(matches[i] = elem);
							}
						}
					}) :
					function( elem, context, xml ) {
						input[0] = elem;
						matcher( input, null, xml, results );
						// Don't keep the element (issue #299)
						input[0] = null;
						return !results.pop();
					};
			}),
	
			"has": markFunction(function( selector ) {
				return function( elem ) {
					return Sizzle( selector, elem ).length > 0;
				};
			}),
	
			"contains": markFunction(function( text ) {
				text = text.replace( runescape, funescape );
				return function( elem ) {
					return ( elem.textContent || elem.innerText || getText( elem ) ).indexOf( text ) > -1;
				};
			}),
	
			// "Whether an element is represented by a :lang() selector
			// is based solely on the element's language value
			// being equal to the identifier C,
			// or beginning with the identifier C immediately followed by "-".
			// The matching of C against the element's language value is performed case-insensitively.
			// The identifier C does not have to be a valid language name."
			// http://www.w3.org/TR/selectors/#lang-pseudo
			"lang": markFunction( function( lang ) {
				// lang value must be a valid identifier
				if ( !ridentifier.test(lang || "") ) {
					Sizzle.error( "unsupported lang: " + lang );
				}
				lang = lang.replace( runescape, funescape ).toLowerCase();
				return function( elem ) {
					var elemLang;
					do {
						if ( (elemLang = documentIsHTML ?
							elem.lang :
							elem.getAttribute("xml:lang") || elem.getAttribute("lang")) ) {
	
							elemLang = elemLang.toLowerCase();
							return elemLang === lang || elemLang.indexOf( lang + "-" ) === 0;
						}
					} while ( (elem = elem.parentNode) && elem.nodeType === 1 );
					return false;
				};
			}),
	
			// Miscellaneous
			"target": function( elem ) {
				var hash = window.location && window.location.hash;
				return hash && hash.slice( 1 ) === elem.id;
			},
	
			"root": function( elem ) {
				return elem === docElem;
			},
	
			"focus": function( elem ) {
				return elem === document.activeElement && (!document.hasFocus || document.hasFocus()) && !!(elem.type || elem.href || ~elem.tabIndex);
			},
	
			// Boolean properties
			"enabled": createDisabledPseudo( false ),
			"disabled": createDisabledPseudo( true ),
	
			"checked": function( elem ) {
				// In CSS3, :checked should return both checked and selected elements
				// http://www.w3.org/TR/2011/REC-css3-selectors-20110929/#checked
				var nodeName = elem.nodeName.toLowerCase();
				return (nodeName === "input" && !!elem.checked) || (nodeName === "option" && !!elem.selected);
			},
	
			"selected": function( elem ) {
				// Accessing this property makes selected-by-default
				// options in Safari work properly
				if ( elem.parentNode ) {
					elem.parentNode.selectedIndex;
				}
	
				return elem.selected === true;
			},
	
			// Contents
			"empty": function( elem ) {
				// http://www.w3.org/TR/selectors/#empty-pseudo
				// :empty is negated by element (1) or content nodes (text: 3; cdata: 4; entity ref: 5),
				//   but not by others (comment: 8; processing instruction: 7; etc.)
				// nodeType < 6 works because attributes (2) do not appear as children
				for ( elem = elem.firstChild; elem; elem = elem.nextSibling ) {
					if ( elem.nodeType < 6 ) {
						return false;
					}
				}
				return true;
			},
	
			"parent": function( elem ) {
				return !Expr.pseudos["empty"]( elem );
			},
	
			// Element/input types
			"header": function( elem ) {
				return rheader.test( elem.nodeName );
			},
	
			"input": function( elem ) {
				return rinputs.test( elem.nodeName );
			},
	
			"button": function( elem ) {
				var name = elem.nodeName.toLowerCase();
				return name === "input" && elem.type === "button" || name === "button";
			},
	
			"text": function( elem ) {
				var attr;
				return elem.nodeName.toLowerCase() === "input" &&
					elem.type === "text" &&
	
					// Support: IE<8
					// New HTML5 attribute values (e.g., "search") appear with elem.type === "text"
					( (attr = elem.getAttribute("type")) == null || attr.toLowerCase() === "text" );
			},
	
			// Position-in-collection
			"first": createPositionalPseudo(function() {
				return [ 0 ];
			}),
	
			"last": createPositionalPseudo(function( matchIndexes, length ) {
				return [ length - 1 ];
			}),
	
			"eq": createPositionalPseudo(function( matchIndexes, length, argument ) {
				return [ argument < 0 ? argument + length : argument ];
			}),
	
			"even": createPositionalPseudo(function( matchIndexes, length ) {
				var i = 0;
				for ( ; i < length; i += 2 ) {
					matchIndexes.push( i );
				}
				return matchIndexes;
			}),
	
			"odd": createPositionalPseudo(function( matchIndexes, length ) {
				var i = 1;
				for ( ; i < length; i += 2 ) {
					matchIndexes.push( i );
				}
				return matchIndexes;
			}),
	
			"lt": createPositionalPseudo(function( matchIndexes, length, argument ) {
				var i = argument < 0 ? argument + length : argument;
				for ( ; --i >= 0; ) {
					matchIndexes.push( i );
				}
				return matchIndexes;
			}),
	
			"gt": createPositionalPseudo(function( matchIndexes, length, argument ) {
				var i = argument < 0 ? argument + length : argument;
				for ( ; ++i < length; ) {
					matchIndexes.push( i );
				}
				return matchIndexes;
			})
		}
	};
	
	Expr.pseudos["nth"] = Expr.pseudos["eq"];
	
	// Add button/input type pseudos
	for ( i in { radio: true, checkbox: true, file: true, password: true, image: true } ) {
		Expr.pseudos[ i ] = createInputPseudo( i );
	}
	for ( i in { submit: true, reset: true } ) {
		Expr.pseudos[ i ] = createButtonPseudo( i );
	}
	
	// Easy API for creating new setFilters
	function setFilters() {}
	setFilters.prototype = Expr.filters = Expr.pseudos;
	Expr.setFilters = new setFilters();
	
	tokenize = Sizzle.tokenize = function( selector, parseOnly ) {
		var matched, match, tokens, type,
			soFar, groups, preFilters,
			cached = tokenCache[ selector + " " ];
	
		if ( cached ) {
			return parseOnly ? 0 : cached.slice( 0 );
		}
	
		soFar = selector;
		groups = [];
		preFilters = Expr.preFilter;
	
		while ( soFar ) {
	
			// Comma and first run
			if ( !matched || (match = rcomma.exec( soFar )) ) {
				if ( match ) {
					// Don't consume trailing commas as valid
					soFar = soFar.slice( match[0].length ) || soFar;
				}
				groups.push( (tokens = []) );
			}
	
			matched = false;
	
			// Combinators
			if ( (match = rcombinators.exec( soFar )) ) {
				matched = match.shift();
				tokens.push({
					value: matched,
					// Cast descendant combinators to space
					type: match[0].replace( rtrim, " " )
				});
				soFar = soFar.slice( matched.length );
			}
	
			// Filters
			for ( type in Expr.filter ) {
				if ( (match = matchExpr[ type ].exec( soFar )) && (!preFilters[ type ] ||
					(match = preFilters[ type ]( match ))) ) {
					matched = match.shift();
					tokens.push({
						value: matched,
						type: type,
						matches: match
					});
					soFar = soFar.slice( matched.length );
				}
			}
	
			if ( !matched ) {
				break;
			}
		}
	
		// Return the length of the invalid excess
		// if we're just parsing
		// Otherwise, throw an error or return tokens
		return parseOnly ?
			soFar.length :
			soFar ?
				Sizzle.error( selector ) :
				// Cache the tokens
				tokenCache( selector, groups ).slice( 0 );
	};
	
	function toSelector( tokens ) {
		var i = 0,
			len = tokens.length,
			selector = "";
		for ( ; i < len; i++ ) {
			selector += tokens[i].value;
		}
		return selector;
	}
	
	function addCombinator( matcher, combinator, base ) {
		var dir = combinator.dir,
			skip = combinator.next,
			key = skip || dir,
			checkNonElements = base && key === "parentNode",
			doneName = done++;
	
		return combinator.first ?
			// Check against closest ancestor/preceding element
			function( elem, context, xml ) {
				while ( (elem = elem[ dir ]) ) {
					if ( elem.nodeType === 1 || checkNonElements ) {
						return matcher( elem, context, xml );
					}
				}
				return false;
			} :
	
			// Check against all ancestor/preceding elements
			function( elem, context, xml ) {
				var oldCache, uniqueCache, outerCache,
					newCache = [ dirruns, doneName ];
	
				// We can't set arbitrary data on XML nodes, so they don't benefit from combinator caching
				if ( xml ) {
					while ( (elem = elem[ dir ]) ) {
						if ( elem.nodeType === 1 || checkNonElements ) {
							if ( matcher( elem, context, xml ) ) {
								return true;
							}
						}
					}
				} else {
					while ( (elem = elem[ dir ]) ) {
						if ( elem.nodeType === 1 || checkNonElements ) {
							outerCache = elem[ expando ] || (elem[ expando ] = {});
	
							// Support: IE <9 only
							// Defend against cloned attroperties (jQuery gh-1709)
							uniqueCache = outerCache[ elem.uniqueID ] || (outerCache[ elem.uniqueID ] = {});
	
							if ( skip && skip === elem.nodeName.toLowerCase() ) {
								elem = elem[ dir ] || elem;
							} else if ( (oldCache = uniqueCache[ key ]) &&
								oldCache[ 0 ] === dirruns && oldCache[ 1 ] === doneName ) {
	
								// Assign to newCache so results back-propagate to previous elements
								return (newCache[ 2 ] = oldCache[ 2 ]);
							} else {
								// Reuse newcache so results back-propagate to previous elements
								uniqueCache[ key ] = newCache;
	
								// A match means we're done; a fail means we have to keep checking
								if ( (newCache[ 2 ] = matcher( elem, context, xml )) ) {
									return true;
								}
							}
						}
					}
				}
				return false;
			};
	}
	
	function elementMatcher( matchers ) {
		return matchers.length > 1 ?
			function( elem, context, xml ) {
				var i = matchers.length;
				while ( i-- ) {
					if ( !matchers[i]( elem, context, xml ) ) {
						return false;
					}
				}
				return true;
			} :
			matchers[0];
	}
	
	function multipleContexts( selector, contexts, results ) {
		var i = 0,
			len = contexts.length;
		for ( ; i < len; i++ ) {
			Sizzle( selector, contexts[i], results );
		}
		return results;
	}
	
	function condense( unmatched, map, filter, context, xml ) {
		var elem,
			newUnmatched = [],
			i = 0,
			len = unmatched.length,
			mapped = map != null;
	
		for ( ; i < len; i++ ) {
			if ( (elem = unmatched[i]) ) {
				if ( !filter || filter( elem, context, xml ) ) {
					newUnmatched.push( elem );
					if ( mapped ) {
						map.push( i );
					}
				}
			}
		}
	
		return newUnmatched;
	}
	
	function setMatcher( preFilter, selector, matcher, postFilter, postFinder, postSelector ) {
		if ( postFilter && !postFilter[ expando ] ) {
			postFilter = setMatcher( postFilter );
		}
		if ( postFinder && !postFinder[ expando ] ) {
			postFinder = setMatcher( postFinder, postSelector );
		}
		return markFunction(function( seed, results, context, xml ) {
			var temp, i, elem,
				preMap = [],
				postMap = [],
				preexisting = results.length,
	
				// Get initial elements from seed or context
				elems = seed || multipleContexts( selector || "*", context.nodeType ? [ context ] : context, [] ),
	
				// Prefilter to get matcher input, preserving a map for seed-results synchronization
				matcherIn = preFilter && ( seed || !selector ) ?
					condense( elems, preMap, preFilter, context, xml ) :
					elems,
	
				matcherOut = matcher ?
					// If we have a postFinder, or filtered seed, or non-seed postFilter or preexisting results,
					postFinder || ( seed ? preFilter : preexisting || postFilter ) ?
	
						// ...intermediate processing is necessary
						[] :
	
						// ...otherwise use results directly
						results :
					matcherIn;
	
			// Find primary matches
			if ( matcher ) {
				matcher( matcherIn, matcherOut, context, xml );
			}
	
			// Apply postFilter
			if ( postFilter ) {
				temp = condense( matcherOut, postMap );
				postFilter( temp, [], context, xml );
	
				// Un-match failing elements by moving them back to matcherIn
				i = temp.length;
				while ( i-- ) {
					if ( (elem = temp[i]) ) {
						matcherOut[ postMap[i] ] = !(matcherIn[ postMap[i] ] = elem);
					}
				}
			}
	
			if ( seed ) {
				if ( postFinder || preFilter ) {
					if ( postFinder ) {
						// Get the final matcherOut by condensing this intermediate into postFinder contexts
						temp = [];
						i = matcherOut.length;
						while ( i-- ) {
							if ( (elem = matcherOut[i]) ) {
								// Restore matcherIn since elem is not yet a final match
								temp.push( (matcherIn[i] = elem) );
							}
						}
						postFinder( null, (matcherOut = []), temp, xml );
					}
	
					// Move matched elements from seed to results to keep them synchronized
					i = matcherOut.length;
					while ( i-- ) {
						if ( (elem = matcherOut[i]) &&
							(temp = postFinder ? indexOf( seed, elem ) : preMap[i]) > -1 ) {
	
							seed[temp] = !(results[temp] = elem);
						}
					}
				}
	
			// Add elements to results, through postFinder if defined
			} else {
				matcherOut = condense(
					matcherOut === results ?
						matcherOut.splice( preexisting, matcherOut.length ) :
						matcherOut
				);
				if ( postFinder ) {
					postFinder( null, results, matcherOut, xml );
				} else {
					push.apply( results, matcherOut );
				}
			}
		});
	}
	
	function matcherFromTokens( tokens ) {
		var checkContext, matcher, j,
			len = tokens.length,
			leadingRelative = Expr.relative[ tokens[0].type ],
			implicitRelative = leadingRelative || Expr.relative[" "],
			i = leadingRelative ? 1 : 0,
	
			// The foundational matcher ensures that elements are reachable from top-level context(s)
			matchContext = addCombinator( function( elem ) {
				return elem === checkContext;
			}, implicitRelative, true ),
			matchAnyContext = addCombinator( function( elem ) {
				return indexOf( checkContext, elem ) > -1;
			}, implicitRelative, true ),
			matchers = [ function( elem, context, xml ) {
				var ret = ( !leadingRelative && ( xml || context !== outermostContext ) ) || (
					(checkContext = context).nodeType ?
						matchContext( elem, context, xml ) :
						matchAnyContext( elem, context, xml ) );
				// Avoid hanging onto element (issue #299)
				checkContext = null;
				return ret;
			} ];
	
		for ( ; i < len; i++ ) {
			if ( (matcher = Expr.relative[ tokens[i].type ]) ) {
				matchers = [ addCombinator(elementMatcher( matchers ), matcher) ];
			} else {
				matcher = Expr.filter[ tokens[i].type ].apply( null, tokens[i].matches );
	
				// Return special upon seeing a positional matcher
				if ( matcher[ expando ] ) {
					// Find the next relative operator (if any) for proper handling
					j = ++i;
					for ( ; j < len; j++ ) {
						if ( Expr.relative[ tokens[j].type ] ) {
							break;
						}
					}
					return setMatcher(
						i > 1 && elementMatcher( matchers ),
						i > 1 && toSelector(
							// If the preceding token was a descendant combinator, insert an implicit any-element `*`
							tokens.slice( 0, i - 1 ).concat({ value: tokens[ i - 2 ].type === " " ? "*" : "" })
						).replace( rtrim, "$1" ),
						matcher,
						i < j && matcherFromTokens( tokens.slice( i, j ) ),
						j < len && matcherFromTokens( (tokens = tokens.slice( j )) ),
						j < len && toSelector( tokens )
					);
				}
				matchers.push( matcher );
			}
		}
	
		return elementMatcher( matchers );
	}
	
	function matcherFromGroupMatchers( elementMatchers, setMatchers ) {
		var bySet = setMatchers.length > 0,
			byElement = elementMatchers.length > 0,
			superMatcher = function( seed, context, xml, results, outermost ) {
				var elem, j, matcher,
					matchedCount = 0,
					i = "0",
					unmatched = seed && [],
					setMatched = [],
					contextBackup = outermostContext,
					// We must always have either seed elements or outermost context
					elems = seed || byElement && Expr.find["TAG"]( "*", outermost ),
					// Use integer dirruns iff this is the outermost matcher
					dirrunsUnique = (dirruns += contextBackup == null ? 1 : Math.random() || 0.1),
					len = elems.length;
	
				if ( outermost ) {
					outermostContext = context === document || context || outermost;
				}
	
				// Add elements passing elementMatchers directly to results
				// Support: IE<9, Safari
				// Tolerate NodeList properties (IE: "length"; Safari: <number>) matching elements by id
				for ( ; i !== len && (elem = elems[i]) != null; i++ ) {
					if ( byElement && elem ) {
						j = 0;
						if ( !context && elem.ownerDocument !== document ) {
							setDocument( elem );
							xml = !documentIsHTML;
						}
						while ( (matcher = elementMatchers[j++]) ) {
							if ( matcher( elem, context || document, xml) ) {
								results.push( elem );
								break;
							}
						}
						if ( outermost ) {
							dirruns = dirrunsUnique;
						}
					}
	
					// Track unmatched elements for set filters
					if ( bySet ) {
						// They will have gone through all possible matchers
						if ( (elem = !matcher && elem) ) {
							matchedCount--;
						}
	
						// Lengthen the array for every element, matched or not
						if ( seed ) {
							unmatched.push( elem );
						}
					}
				}
	
				// `i` is now the count of elements visited above, and adding it to `matchedCount`
				// makes the latter nonnegative.
				matchedCount += i;
	
				// Apply set filters to unmatched elements
				// NOTE: This can be skipped if there are no unmatched elements (i.e., `matchedCount`
				// equals `i`), unless we didn't visit _any_ elements in the above loop because we have
				// no element matchers and no seed.
				// Incrementing an initially-string "0" `i` allows `i` to remain a string only in that
				// case, which will result in a "00" `matchedCount` that differs from `i` but is also
				// numerically zero.
				if ( bySet && i !== matchedCount ) {
					j = 0;
					while ( (matcher = setMatchers[j++]) ) {
						matcher( unmatched, setMatched, context, xml );
					}
	
					if ( seed ) {
						// Reintegrate element matches to eliminate the need for sorting
						if ( matchedCount > 0 ) {
							while ( i-- ) {
								if ( !(unmatched[i] || setMatched[i]) ) {
									setMatched[i] = pop.call( results );
								}
							}
						}
	
						// Discard index placeholder values to get only actual matches
						setMatched = condense( setMatched );
					}
	
					// Add matches to results
					push.apply( results, setMatched );
	
					// Seedless set matches succeeding multiple successful matchers stipulate sorting
					if ( outermost && !seed && setMatched.length > 0 &&
						( matchedCount + setMatchers.length ) > 1 ) {
	
						Sizzle.uniqueSort( results );
					}
				}
	
				// Override manipulation of globals by nested matchers
				if ( outermost ) {
					dirruns = dirrunsUnique;
					outermostContext = contextBackup;
				}
	
				return unmatched;
			};
	
		return bySet ?
			markFunction( superMatcher ) :
			superMatcher;
	}
	
	compile = Sizzle.compile = function( selector, match /* Internal Use Only */ ) {
		var i,
			setMatchers = [],
			elementMatchers = [],
			cached = compilerCache[ selector + " " ];
	
		if ( !cached ) {
			// Generate a function of recursive functions that can be used to check each element
			if ( !match ) {
				match = tokenize( selector );
			}
			i = match.length;
			while ( i-- ) {
				cached = matcherFromTokens( match[i] );
				if ( cached[ expando ] ) {
					setMatchers.push( cached );
				} else {
					elementMatchers.push( cached );
				}
			}
	
			// Cache the compiled function
			cached = compilerCache( selector, matcherFromGroupMatchers( elementMatchers, setMatchers ) );
	
			// Save selector and tokenization
			cached.selector = selector;
		}
		return cached;
	};
	
	/**
	 * A low-level selection function that works with Sizzle's compiled
	 *  selector functions
	 * @param {String|Function} selector A selector or a pre-compiled
	 *  selector function built with Sizzle.compile
	 * @param {Element} context
	 * @param {Array} [results]
	 * @param {Array} [seed] A set of elements to match against
	 */
	select = Sizzle.select = function( selector, context, results, seed ) {
		var i, tokens, token, type, find,
			compiled = typeof selector === "function" && selector,
			match = !seed && tokenize( (selector = compiled.selector || selector) );
	
		results = results || [];
	
		// Try to minimize operations if there is only one selector in the list and no seed
		// (the latter of which guarantees us context)
		if ( match.length === 1 ) {
	
			// Reduce context if the leading compound selector is an ID
			tokens = match[0] = match[0].slice( 0 );
			if ( tokens.length > 2 && (token = tokens[0]).type === "ID" &&
					context.nodeType === 9 && documentIsHTML && Expr.relative[ tokens[1].type ] ) {
	
				context = ( Expr.find["ID"]( token.matches[0].replace(runescape, funescape), context ) || [] )[0];
				if ( !context ) {
					return results;
	
				// Precompiled matchers will still verify ancestry, so step up a level
				} else if ( compiled ) {
					context = context.parentNode;
				}
	
				selector = selector.slice( tokens.shift().value.length );
			}
	
			// Fetch a seed set for right-to-left matching
			i = matchExpr["needsContext"].test( selector ) ? 0 : tokens.length;
			while ( i-- ) {
				token = tokens[i];
	
				// Abort if we hit a combinator
				if ( Expr.relative[ (type = token.type) ] ) {
					break;
				}
				if ( (find = Expr.find[ type ]) ) {
					// Search, expanding context for leading sibling combinators
					if ( (seed = find(
						token.matches[0].replace( runescape, funescape ),
						rsibling.test( tokens[0].type ) && testContext( context.parentNode ) || context
					)) ) {
	
						// If seed is empty or no tokens remain, we can return early
						tokens.splice( i, 1 );
						selector = seed.length && toSelector( tokens );
						if ( !selector ) {
							push.apply( results, seed );
							return results;
						}
	
						break;
					}
				}
			}
		}
	
		// Compile and execute a filtering function if one is not provided
		// Provide `match` to avoid retokenization if we modified the selector above
		( compiled || compile( selector, match ) )(
			seed,
			context,
			!documentIsHTML,
			results,
			!context || rsibling.test( selector ) && testContext( context.parentNode ) || context
		);
		return results;
	};
	
	// One-time assignments
	
	// Sort stability
	support.sortStable = expando.split("").sort( sortOrder ).join("") === expando;
	
	// Support: Chrome 14-35+
	// Always assume duplicates if they aren't passed to the comparison function
	support.detectDuplicates = !!hasDuplicate;
	
	// Initialize against the default document
	setDocument();
	
	// Support: Webkit<537.32 - Safari 6.0.3/Chrome 25 (fixed in Chrome 27)
	// Detached nodes confoundingly follow *each other*
	support.sortDetached = assert(function( el ) {
		// Should return 1, but returns 4 (following)
		return el.compareDocumentPosition( document.createElement("fieldset") ) & 1;
	});
	
	// Support: IE<8
	// Prevent attribute/property "interpolation"
	// https://msdn.microsoft.com/en-us/library/ms536429%28VS.85%29.aspx
	if ( !assert(function( el ) {
		el.innerHTML = "<a href='#'></a>";
		return el.firstChild.getAttribute("href") === "#" ;
	}) ) {
		addHandle( "type|href|height|width", function( elem, name, isXML ) {
			if ( !isXML ) {
				return elem.getAttribute( name, name.toLowerCase() === "type" ? 1 : 2 );
			}
		});
	}
	
	// Support: IE<9
	// Use defaultValue in place of getAttribute("value")
	if ( !support.attributes || !assert(function( el ) {
		el.innerHTML = "<input/>";
		el.firstChild.setAttribute( "value", "" );
		return el.firstChild.getAttribute( "value" ) === "";
	}) ) {
		addHandle( "value", function( elem, name, isXML ) {
			if ( !isXML && elem.nodeName.toLowerCase() === "input" ) {
				return elem.defaultValue;
			}
		});
	}
	
	// Support: IE<9
	// Use getAttributeNode to fetch booleans when getAttribute lies
	if ( !assert(function( el ) {
		return el.getAttribute("disabled") == null;
	}) ) {
		addHandle( booleans, function( elem, name, isXML ) {
			var val;
			if ( !isXML ) {
				return elem[ name ] === true ? name.toLowerCase() :
						(val = elem.getAttributeNode( name )) && val.specified ?
						val.value :
					null;
			}
		});
	}
	
	return Sizzle;
	
	})( window );
	
	
	
	jQuery.find = Sizzle;
	jQuery.expr = Sizzle.selectors;
	
	// Deprecated
	jQuery.expr[ ":" ] = jQuery.expr.pseudos;
	jQuery.uniqueSort = jQuery.unique = Sizzle.uniqueSort;
	jQuery.text = Sizzle.getText;
	jQuery.isXMLDoc = Sizzle.isXML;
	jQuery.contains = Sizzle.contains;
	jQuery.escapeSelector = Sizzle.escape;
	
	
	
	
	var dir = function( elem, dir, until ) {
		var matched = [],
			truncate = until !== undefined;
	
		while ( ( elem = elem[ dir ] ) && elem.nodeType !== 9 ) {
			if ( elem.nodeType === 1 ) {
				if ( truncate && jQuery( elem ).is( until ) ) {
					break;
				}
				matched.push( elem );
			}
		}
		return matched;
	};
	
	
	var siblings = function( n, elem ) {
		var matched = [];
	
		for ( ; n; n = n.nextSibling ) {
			if ( n.nodeType === 1 && n !== elem ) {
				matched.push( n );
			}
		}
	
		return matched;
	};
	
	
	var rneedsContext = jQuery.expr.match.needsContext;
	
	
	
	function nodeName( elem, name ) {
	
	  return elem.nodeName && elem.nodeName.toLowerCase() === name.toLowerCase();
	
	};
	var rsingleTag = ( /^<([a-z][^\/\0>:\x20\t\r\n\f]*)[\x20\t\r\n\f]*\/?>(?:<\/\1>|)$/i );
	
	
	
	var risSimple = /^.[^:#\[\.,]*$/;
	
	// Implement the identical functionality for filter and not
	function winnow( elements, qualifier, not ) {
		if ( jQuery.isFunction( qualifier ) ) {
			return jQuery.grep( elements, function( elem, i ) {
				return !!qualifier.call( elem, i, elem ) !== not;
			} );
		}
	
		// Single element
		if ( qualifier.nodeType ) {
			return jQuery.grep( elements, function( elem ) {
				return ( elem === qualifier ) !== not;
			} );
		}
	
		// Arraylike of elements (jQuery, arguments, Array)
		if ( typeof qualifier !== "string" ) {
			return jQuery.grep( elements, function( elem ) {
				return ( indexOf.call( qualifier, elem ) > -1 ) !== not;
			} );
		}
	
		// Simple selector that can be filtered directly, removing non-Elements
		if ( risSimple.test( qualifier ) ) {
			return jQuery.filter( qualifier, elements, not );
		}
	
		// Complex selector, compare the two sets, removing non-Elements
		qualifier = jQuery.filter( qualifier, elements );
		return jQuery.grep( elements, function( elem ) {
			return ( indexOf.call( qualifier, elem ) > -1 ) !== not && elem.nodeType === 1;
		} );
	}
	
	jQuery.filter = function( expr, elems, not ) {
		var elem = elems[ 0 ];
	
		if ( not ) {
			expr = ":not(" + expr + ")";
		}
	
		if ( elems.length === 1 && elem.nodeType === 1 ) {
			return jQuery.find.matchesSelector( elem, expr ) ? [ elem ] : [];
		}
	
		return jQuery.find.matches( expr, jQuery.grep( elems, function( elem ) {
			return elem.nodeType === 1;
		} ) );
	};
	
	jQuery.fn.extend( {
		find: function( selector ) {
			var i, ret,
				len = this.length,
				self = this;
	
			if ( typeof selector !== "string" ) {
				return this.pushStack( jQuery( selector ).filter( function() {
					for ( i = 0; i < len; i++ ) {
						if ( jQuery.contains( self[ i ], this ) ) {
							return true;
						}
					}
				} ) );
			}
	
			ret = this.pushStack( [] );
	
			for ( i = 0; i < len; i++ ) {
				jQuery.find( selector, self[ i ], ret );
			}
	
			return len > 1 ? jQuery.uniqueSort( ret ) : ret;
		},
		filter: function( selector ) {
			return this.pushStack( winnow( this, selector || [], false ) );
		},
		not: function( selector ) {
			return this.pushStack( winnow( this, selector || [], true ) );
		},
		is: function( selector ) {
			return !!winnow(
				this,
	
				// If this is a positional/relative selector, check membership in the returned set
				// so $("p:first").is("p:last") won't return true for a doc with two "p".
				typeof selector === "string" && rneedsContext.test( selector ) ?
					jQuery( selector ) :
					selector || [],
				false
			).length;
		}
	} );
	
	
	// Initialize a jQuery object
	
	
	// A central reference to the root jQuery(document)
	var rootjQuery,
	
		// A simple way to check for HTML strings
		// Prioritize #id over <tag> to avoid XSS via location.hash (#9521)
		// Strict HTML recognition (#11290: must start with <)
		// Shortcut simple #id case for speed
		rquickExpr = /^(?:\s*(<[\w\W]+>)[^>]*|#([\w-]+))$/,
	
		init = jQuery.fn.init = function( selector, context, root ) {
			var match, elem;
	
			// HANDLE: $(""), $(null), $(undefined), $(false)
			if ( !selector ) {
				return this;
			}
	
			// Method init() accepts an alternate rootjQuery
			// so migrate can support jQuery.sub (gh-2101)
			root = root || rootjQuery;
	
			// Handle HTML strings
			if ( typeof selector === "string" ) {
				if ( selector[ 0 ] === "<" &&
					selector[ selector.length - 1 ] === ">" &&
					selector.length >= 3 ) {
	
					// Assume that strings that start and end with <> are HTML and skip the regex check
					match = [ null, selector, null ];
	
				} else {
					match = rquickExpr.exec( selector );
				}
	
				// Match html or make sure no context is specified for #id
				if ( match && ( match[ 1 ] || !context ) ) {
	
					// HANDLE: $(html) -> $(array)
					if ( match[ 1 ] ) {
						context = context instanceof jQuery ? context[ 0 ] : context;
	
						// Option to run scripts is true for back-compat
						// Intentionally let the error be thrown if parseHTML is not present
						jQuery.merge( this, jQuery.parseHTML(
							match[ 1 ],
							context && context.nodeType ? context.ownerDocument || context : document,
							true
						) );
	
						// HANDLE: $(html, props)
						if ( rsingleTag.test( match[ 1 ] ) && jQuery.isPlainObject( context ) ) {
							for ( match in context ) {
	
								// Properties of context are called as methods if possible
								if ( jQuery.isFunction( this[ match ] ) ) {
									this[ match ]( context[ match ] );
	
								// ...and otherwise set as attributes
								} else {
									this.attr( match, context[ match ] );
								}
							}
						}
	
						return this;
	
					// HANDLE: $(#id)
					} else {
						elem = document.getElementById( match[ 2 ] );
	
						if ( elem ) {
	
							// Inject the element directly into the jQuery object
							this[ 0 ] = elem;
							this.length = 1;
						}
						return this;
					}
	
				// HANDLE: $(expr, $(...))
				} else if ( !context || context.jquery ) {
					return ( context || root ).find( selector );
	
				// HANDLE: $(expr, context)
				// (which is just equivalent to: $(context).find(expr)
				} else {
					return this.constructor( context ).find( selector );
				}
	
			// HANDLE: $(DOMElement)
			} else if ( selector.nodeType ) {
				this[ 0 ] = selector;
				this.length = 1;
				return this;
	
			// HANDLE: $(function)
			// Shortcut for document ready
			} else if ( jQuery.isFunction( selector ) ) {
				return root.ready !== undefined ?
					root.ready( selector ) :
	
					// Execute immediately if ready is not present
					selector( jQuery );
			}
	
			return jQuery.makeArray( selector, this );
		};
	
	// Give the init function the jQuery prototype for later instantiation
	init.prototype = jQuery.fn;
	
	// Initialize central reference
	rootjQuery = jQuery( document );
	
	
	var rparentsprev = /^(?:parents|prev(?:Until|All))/,
	
		// Methods guaranteed to produce a unique set when starting from a unique set
		guaranteedUnique = {
			children: true,
			contents: true,
			next: true,
			prev: true
		};
	
	jQuery.fn.extend( {
		has: function( target ) {
			var targets = jQuery( target, this ),
				l = targets.length;
	
			return this.filter( function() {
				var i = 0;
				for ( ; i < l; i++ ) {
					if ( jQuery.contains( this, targets[ i ] ) ) {
						return true;
					}
				}
			} );
		},
	
		closest: function( selectors, context ) {
			var cur,
				i = 0,
				l = this.length,
				matched = [],
				targets = typeof selectors !== "string" && jQuery( selectors );
	
			// Positional selectors never match, since there's no _selection_ context
			if ( !rneedsContext.test( selectors ) ) {
				for ( ; i < l; i++ ) {
					for ( cur = this[ i ]; cur && cur !== context; cur = cur.parentNode ) {
	
						// Always skip document fragments
						if ( cur.nodeType < 11 && ( targets ?
							targets.index( cur ) > -1 :
	
							// Don't pass non-elements to Sizzle
							cur.nodeType === 1 &&
								jQuery.find.matchesSelector( cur, selectors ) ) ) {
	
							matched.push( cur );
							break;
						}
					}
				}
			}
	
			return this.pushStack( matched.length > 1 ? jQuery.uniqueSort( matched ) : matched );
		},
	
		// Determine the position of an element within the set
		index: function( elem ) {
	
			// No argument, return index in parent
			if ( !elem ) {
				return ( this[ 0 ] && this[ 0 ].parentNode ) ? this.first().prevAll().length : -1;
			}
	
			// Index in selector
			if ( typeof elem === "string" ) {
				return indexOf.call( jQuery( elem ), this[ 0 ] );
			}
	
			// Locate the position of the desired element
			return indexOf.call( this,
	
				// If it receives a jQuery object, the first element is used
				elem.jquery ? elem[ 0 ] : elem
			);
		},
	
		add: function( selector, context ) {
			return this.pushStack(
				jQuery.uniqueSort(
					jQuery.merge( this.get(), jQuery( selector, context ) )
				)
			);
		},
	
		addBack: function( selector ) {
			return this.add( selector == null ?
				this.prevObject : this.prevObject.filter( selector )
			);
		}
	} );
	
	function sibling( cur, dir ) {
		while ( ( cur = cur[ dir ] ) && cur.nodeType !== 1 ) {}
		return cur;
	}
	
	jQuery.each( {
		parent: function( elem ) {
			var parent = elem.parentNode;
			return parent && parent.nodeType !== 11 ? parent : null;
		},
		parents: function( elem ) {
			return dir( elem, "parentNode" );
		},
		parentsUntil: function( elem, i, until ) {
			return dir( elem, "parentNode", until );
		},
		next: function( elem ) {
			return sibling( elem, "nextSibling" );
		},
		prev: function( elem ) {
			return sibling( elem, "previousSibling" );
		},
		nextAll: function( elem ) {
			return dir( elem, "nextSibling" );
		},
		prevAll: function( elem ) {
			return dir( elem, "previousSibling" );
		},
		nextUntil: function( elem, i, until ) {
			return dir( elem, "nextSibling", until );
		},
		prevUntil: function( elem, i, until ) {
			return dir( elem, "previousSibling", until );
		},
		siblings: function( elem ) {
			return siblings( ( elem.parentNode || {} ).firstChild, elem );
		},
		children: function( elem ) {
			return siblings( elem.firstChild );
		},
		contents: function( elem ) {
	        if ( nodeName( elem, "iframe" ) ) {
	            return elem.contentDocument;
	        }
	
	        // Support: IE 9 - 11 only, iOS 7 only, Android Browser <=4.3 only
	        // Treat the template element as a regular one in browsers that
	        // don't support it.
	        if ( nodeName( elem, "template" ) ) {
	            elem = elem.content || elem;
	        }
	
	        return jQuery.merge( [], elem.childNodes );
		}
	}, function( name, fn ) {
		jQuery.fn[ name ] = function( until, selector ) {
			var matched = jQuery.map( this, fn, until );
	
			if ( name.slice( -5 ) !== "Until" ) {
				selector = until;
			}
	
			if ( selector && typeof selector === "string" ) {
				matched = jQuery.filter( selector, matched );
			}
	
			if ( this.length > 1 ) {
	
				// Remove duplicates
				if ( !guaranteedUnique[ name ] ) {
					jQuery.uniqueSort( matched );
				}
	
				// Reverse order for parents* and prev-derivatives
				if ( rparentsprev.test( name ) ) {
					matched.reverse();
				}
			}
	
			return this.pushStack( matched );
		};
	} );
	var rnothtmlwhite = ( /[^\x20\t\r\n\f]+/g );
	
	
	
	// Convert String-formatted options into Object-formatted ones
	function createOptions( options ) {
		var object = {};
		jQuery.each( options.match( rnothtmlwhite ) || [], function( _, flag ) {
			object[ flag ] = true;
		} );
		return object;
	}
	
	/*
	 * Create a callback list using the following parameters:
	 *
	 *	options: an optional list of space-separated options that will change how
	 *			the callback list behaves or a more traditional option object
	 *
	 * By default a callback list will act like an event callback list and can be
	 * "fired" multiple times.
	 *
	 * Possible options:
	 *
	 *	once:			will ensure the callback list can only be fired once (like a Deferred)
	 *
	 *	memory:			will keep track of previous values and will call any callback added
	 *					after the list has been fired right away with the latest "memorized"
	 *					values (like a Deferred)
	 *
	 *	unique:			will ensure a callback can only be added once (no duplicate in the list)
	 *
	 *	stopOnFalse:	interrupt callings when a callback returns false
	 *
	 */
	jQuery.Callbacks = function( options ) {
	
		// Convert options from String-formatted to Object-formatted if needed
		// (we check in cache first)
		options = typeof options === "string" ?
			createOptions( options ) :
			jQuery.extend( {}, options );
	
		var // Flag to know if list is currently firing
			firing,
	
			// Last fire value for non-forgettable lists
			memory,
	
			// Flag to know if list was already fired
			fired,
	
			// Flag to prevent firing
			locked,
	
			// Actual callback list
			list = [],
	
			// Queue of execution data for repeatable lists
			queue = [],
	
			// Index of currently firing callback (modified by add/remove as needed)
			firingIndex = -1,
	
			// Fire callbacks
			fire = function() {
	
				// Enforce single-firing
				locked = locked || options.once;
	
				// Execute callbacks for all pending executions,
				// respecting firingIndex overrides and runtime changes
				fired = firing = true;
				for ( ; queue.length; firingIndex = -1 ) {
					memory = queue.shift();
					while ( ++firingIndex < list.length ) {
	
						// Run callback and check for early termination
						if ( list[ firingIndex ].apply( memory[ 0 ], memory[ 1 ] ) === false &&
							options.stopOnFalse ) {
	
							// Jump to end and forget the data so .add doesn't re-fire
							firingIndex = list.length;
							memory = false;
						}
					}
				}
	
				// Forget the data if we're done with it
				if ( !options.memory ) {
					memory = false;
				}
	
				firing = false;
	
				// Clean up if we're done firing for good
				if ( locked ) {
	
					// Keep an empty list if we have data for future add calls
					if ( memory ) {
						list = [];
	
					// Otherwise, this object is spent
					} else {
						list = "";
					}
				}
			},
	
			// Actual Callbacks object
			self = {
	
				// Add a callback or a collection of callbacks to the list
				add: function() {
					if ( list ) {
	
						// If we have memory from a past run, we should fire after adding
						if ( memory && !firing ) {
							firingIndex = list.length - 1;
							queue.push( memory );
						}
	
						( function add( args ) {
							jQuery.each( args, function( _, arg ) {
								if ( jQuery.isFunction( arg ) ) {
									if ( !options.unique || !self.has( arg ) ) {
										list.push( arg );
									}
								} else if ( arg && arg.length && jQuery.type( arg ) !== "string" ) {
	
									// Inspect recursively
									add( arg );
								}
							} );
						} )( arguments );
	
						if ( memory && !firing ) {
							fire();
						}
					}
					return this;
				},
	
				// Remove a callback from the list
				remove: function() {
					jQuery.each( arguments, function( _, arg ) {
						var index;
						while ( ( index = jQuery.inArray( arg, list, index ) ) > -1 ) {
							list.splice( index, 1 );
	
							// Handle firing indexes
							if ( index <= firingIndex ) {
								firingIndex--;
							}
						}
					} );
					return this;
				},
	
				// Check if a given callback is in the list.
				// If no argument is given, return whether or not list has callbacks attached.
				has: function( fn ) {
					return fn ?
						jQuery.inArray( fn, list ) > -1 :
						list.length > 0;
				},
	
				// Remove all callbacks from the list
				empty: function() {
					if ( list ) {
						list = [];
					}
					return this;
				},
	
				// Disable .fire and .add
				// Abort any current/pending executions
				// Clear all callbacks and values
				disable: function() {
					locked = queue = [];
					list = memory = "";
					return this;
				},
				disabled: function() {
					return !list;
				},
	
				// Disable .fire
				// Also disable .add unless we have memory (since it would have no effect)
				// Abort any pending executions
				lock: function() {
					locked = queue = [];
					if ( !memory && !firing ) {
						list = memory = "";
					}
					return this;
				},
				locked: function() {
					return !!locked;
				},
	
				// Call all callbacks with the given context and arguments
				fireWith: function( context, args ) {
					if ( !locked ) {
						args = args || [];
						args = [ context, args.slice ? args.slice() : args ];
						queue.push( args );
						if ( !firing ) {
							fire();
						}
					}
					return this;
				},
	
				// Call all the callbacks with the given arguments
				fire: function() {
					self.fireWith( this, arguments );
					return this;
				},
	
				// To know if the callbacks have already been called at least once
				fired: function() {
					return !!fired;
				}
			};
	
		return self;
	};
	
	
	function Identity( v ) {
		return v;
	}
	function Thrower( ex ) {
		throw ex;
	}
	
	function adoptValue( value, resolve, reject, noValue ) {
		var method;
	
		try {
	
			// Check for promise aspect first to privilege synchronous behavior
			if ( value && jQuery.isFunction( ( method = value.promise ) ) ) {
				method.call( value ).done( resolve ).fail( reject );
	
			// Other thenables
			} else if ( value && jQuery.isFunction( ( method = value.then ) ) ) {
				method.call( value, resolve, reject );
	
			// Other non-thenables
			} else {
	
				// Control `resolve` arguments by letting Array#slice cast boolean `noValue` to integer:
				// * false: [ value ].slice( 0 ) => resolve( value )
				// * true: [ value ].slice( 1 ) => resolve()
				resolve.apply( undefined, [ value ].slice( noValue ) );
			}
	
		// For Promises/A+, convert exceptions into rejections
		// Since jQuery.when doesn't unwrap thenables, we can skip the extra checks appearing in
		// Deferred#then to conditionally suppress rejection.
		} catch ( value ) {
	
			// Support: Android 4.0 only
			// Strict mode functions invoked without .call/.apply get global-object context
			reject.apply( undefined, [ value ] );
		}
	}
	
	jQuery.extend( {
	
		Deferred: function( func ) {
			var tuples = [
	
					// action, add listener, callbacks,
					// ... .then handlers, argument index, [final state]
					[ "notify", "progress", jQuery.Callbacks( "memory" ),
						jQuery.Callbacks( "memory" ), 2 ],
					[ "resolve", "done", jQuery.Callbacks( "once memory" ),
						jQuery.Callbacks( "once memory" ), 0, "resolved" ],
					[ "reject", "fail", jQuery.Callbacks( "once memory" ),
						jQuery.Callbacks( "once memory" ), 1, "rejected" ]
				],
				state = "pending",
				promise = {
					state: function() {
						return state;
					},
					always: function() {
						deferred.done( arguments ).fail( arguments );
						return this;
					},
					"catch": function( fn ) {
						return promise.then( null, fn );
					},
	
					// Keep pipe for back-compat
					pipe: function( /* fnDone, fnFail, fnProgress */ ) {
						var fns = arguments;
	
						return jQuery.Deferred( function( newDefer ) {
							jQuery.each( tuples, function( i, tuple ) {
	
								// Map tuples (progress, done, fail) to arguments (done, fail, progress)
								var fn = jQuery.isFunction( fns[ tuple[ 4 ] ] ) && fns[ tuple[ 4 ] ];
	
								// deferred.progress(function() { bind to newDefer or newDefer.notify })
								// deferred.done(function() { bind to newDefer or newDefer.resolve })
								// deferred.fail(function() { bind to newDefer or newDefer.reject })
								deferred[ tuple[ 1 ] ]( function() {
									var returned = fn && fn.apply( this, arguments );
									if ( returned && jQuery.isFunction( returned.promise ) ) {
										returned.promise()
											.progress( newDefer.notify )
											.done( newDefer.resolve )
											.fail( newDefer.reject );
									} else {
										newDefer[ tuple[ 0 ] + "With" ](
											this,
											fn ? [ returned ] : arguments
										);
									}
								} );
							} );
							fns = null;
						} ).promise();
					},
					then: function( onFulfilled, onRejected, onProgress ) {
						var maxDepth = 0;
						function resolve( depth, deferred, handler, special ) {
							return function() {
								var that = this,
									args = arguments,
									mightThrow = function() {
										var returned, then;
	
										// Support: Promises/A+ section 2.3.3.3.3
										// https://promisesaplus.com/#point-59
										// Ignore double-resolution attempts
										if ( depth < maxDepth ) {
											return;
										}
	
										returned = handler.apply( that, args );
	
										// Support: Promises/A+ section 2.3.1
										// https://promisesaplus.com/#point-48
										if ( returned === deferred.promise() ) {
											throw new TypeError( "Thenable self-resolution" );
										}
	
										// Support: Promises/A+ sections 2.3.3.1, 3.5
										// https://promisesaplus.com/#point-54
										// https://promisesaplus.com/#point-75
										// Retrieve `then` only once
										then = returned &&
	
											// Support: Promises/A+ section 2.3.4
											// https://promisesaplus.com/#point-64
											// Only check objects and functions for thenability
											( typeof returned === "object" ||
												typeof returned === "function" ) &&
											returned.then;
	
										// Handle a returned thenable
										if ( jQuery.isFunction( then ) ) {
	
											// Special processors (notify) just wait for resolution
											if ( special ) {
												then.call(
													returned,
													resolve( maxDepth, deferred, Identity, special ),
													resolve( maxDepth, deferred, Thrower, special )
												);
	
											// Normal processors (resolve) also hook into progress
											} else {
	
												// ...and disregard older resolution values
												maxDepth++;
	
												then.call(
													returned,
													resolve( maxDepth, deferred, Identity, special ),
													resolve( maxDepth, deferred, Thrower, special ),
													resolve( maxDepth, deferred, Identity,
														deferred.notifyWith )
												);
											}
	
										// Handle all other returned values
										} else {
	
											// Only substitute handlers pass on context
											// and multiple values (non-spec behavior)
											if ( handler !== Identity ) {
												that = undefined;
												args = [ returned ];
											}
	
											// Process the value(s)
											// Default process is resolve
											( special || deferred.resolveWith )( that, args );
										}
									},
	
									// Only normal processors (resolve) catch and reject exceptions
									process = special ?
										mightThrow :
										function() {
											try {
												mightThrow();
											} catch ( e ) {
	
												if ( jQuery.Deferred.exceptionHook ) {
													jQuery.Deferred.exceptionHook( e,
														process.stackTrace );
												}
	
												// Support: Promises/A+ section 2.3.3.3.4.1
												// https://promisesaplus.com/#point-61
												// Ignore post-resolution exceptions
												if ( depth + 1 >= maxDepth ) {
	
													// Only substitute handlers pass on context
													// and multiple values (non-spec behavior)
													if ( handler !== Thrower ) {
														that = undefined;
														args = [ e ];
													}
	
													deferred.rejectWith( that, args );
												}
											}
										};
	
								// Support: Promises/A+ section 2.3.3.3.1
								// https://promisesaplus.com/#point-57
								// Re-resolve promises immediately to dodge false rejection from
								// subsequent errors
								if ( depth ) {
									process();
								} else {
	
									// Call an optional hook to record the stack, in case of exception
									// since it's otherwise lost when execution goes async
									if ( jQuery.Deferred.getStackHook ) {
										process.stackTrace = jQuery.Deferred.getStackHook();
									}
									window.setTimeout( process );
								}
							};
						}
	
						return jQuery.Deferred( function( newDefer ) {
	
							// progress_handlers.add( ... )
							tuples[ 0 ][ 3 ].add(
								resolve(
									0,
									newDefer,
									jQuery.isFunction( onProgress ) ?
										onProgress :
										Identity,
									newDefer.notifyWith
								)
							);
	
							// fulfilled_handlers.add( ... )
							tuples[ 1 ][ 3 ].add(
								resolve(
									0,
									newDefer,
									jQuery.isFunction( onFulfilled ) ?
										onFulfilled :
										Identity
								)
							);
	
							// rejected_handlers.add( ... )
							tuples[ 2 ][ 3 ].add(
								resolve(
									0,
									newDefer,
									jQuery.isFunction( onRejected ) ?
										onRejected :
										Thrower
								)
							);
						} ).promise();
					},
	
					// Get a promise for this deferred
					// If obj is provided, the promise aspect is added to the object
					promise: function( obj ) {
						return obj != null ? jQuery.extend( obj, promise ) : promise;
					}
				},
				deferred = {};
	
			// Add list-specific methods
			jQuery.each( tuples, function( i, tuple ) {
				var list = tuple[ 2 ],
					stateString = tuple[ 5 ];
	
				// promise.progress = list.add
				// promise.done = list.add
				// promise.fail = list.add
				promise[ tuple[ 1 ] ] = list.add;
	
				// Handle state
				if ( stateString ) {
					list.add(
						function() {
	
							// state = "resolved" (i.e., fulfilled)
							// state = "rejected"
							state = stateString;
						},
	
						// rejected_callbacks.disable
						// fulfilled_callbacks.disable
						tuples[ 3 - i ][ 2 ].disable,
	
						// progress_callbacks.lock
						tuples[ 0 ][ 2 ].lock
					);
				}
	
				// progress_handlers.fire
				// fulfilled_handlers.fire
				// rejected_handlers.fire
				list.add( tuple[ 3 ].fire );
	
				// deferred.notify = function() { deferred.notifyWith(...) }
				// deferred.resolve = function() { deferred.resolveWith(...) }
				// deferred.reject = function() { deferred.rejectWith(...) }
				deferred[ tuple[ 0 ] ] = function() {
					deferred[ tuple[ 0 ] + "With" ]( this === deferred ? undefined : this, arguments );
					return this;
				};
	
				// deferred.notifyWith = list.fireWith
				// deferred.resolveWith = list.fireWith
				// deferred.rejectWith = list.fireWith
				deferred[ tuple[ 0 ] + "With" ] = list.fireWith;
			} );
	
			// Make the deferred a promise
			promise.promise( deferred );
	
			// Call given func if any
			if ( func ) {
				func.call( deferred, deferred );
			}
	
			// All done!
			return deferred;
		},
	
		// Deferred helper
		when: function( singleValue ) {
			var
	
				// count of uncompleted subordinates
				remaining = arguments.length,
	
				// count of unprocessed arguments
				i = remaining,
	
				// subordinate fulfillment data
				resolveContexts = Array( i ),
				resolveValues = slice.call( arguments ),
	
				// the master Deferred
				master = jQuery.Deferred(),
	
				// subordinate callback factory
				updateFunc = function( i ) {
					return function( value ) {
						resolveContexts[ i ] = this;
						resolveValues[ i ] = arguments.length > 1 ? slice.call( arguments ) : value;
						if ( !( --remaining ) ) {
							master.resolveWith( resolveContexts, resolveValues );
						}
					};
				};
	
			// Single- and empty arguments are adopted like Promise.resolve
			if ( remaining <= 1 ) {
				adoptValue( singleValue, master.done( updateFunc( i ) ).resolve, master.reject,
					!remaining );
	
				// Use .then() to unwrap secondary thenables (cf. gh-3000)
				if ( master.state() === "pending" ||
					jQuery.isFunction( resolveValues[ i ] && resolveValues[ i ].then ) ) {
	
					return master.then();
				}
			}
	
			// Multiple arguments are aggregated like Promise.all array elements
			while ( i-- ) {
				adoptValue( resolveValues[ i ], updateFunc( i ), master.reject );
			}
	
			return master.promise();
		}
	} );
	
	
	// These usually indicate a programmer mistake during development,
	// warn about them ASAP rather than swallowing them by default.
	var rerrorNames = /^(Eval|Internal|Range|Reference|Syntax|Type|URI)Error$/;
	
	jQuery.Deferred.exceptionHook = function( error, stack ) {
	
		// Support: IE 8 - 9 only
		// Console exists when dev tools are open, which can happen at any time
		if ( window.console && window.console.warn && error && rerrorNames.test( error.name ) ) {
			window.console.warn( "jQuery.Deferred exception: " + error.message, error.stack, stack );
		}
	};
	
	
	
	
	jQuery.readyException = function( error ) {
		window.setTimeout( function() {
			throw error;
		} );
	};
	
	
	
	
	// The deferred used on DOM ready
	var readyList = jQuery.Deferred();
	
	jQuery.fn.ready = function( fn ) {
	
		readyList
			.then( fn )
	
			// Wrap jQuery.readyException in a function so that the lookup
			// happens at the time of error handling instead of callback
			// registration.
			.catch( function( error ) {
				jQuery.readyException( error );
			} );
	
		return this;
	};
	
	jQuery.extend( {
	
		// Is the DOM ready to be used? Set to true once it occurs.
		isReady: false,
	
		// A counter to track how many items to wait for before
		// the ready event fires. See #6781
		readyWait: 1,
	
		// Handle when the DOM is ready
		ready: function( wait ) {
	
			// Abort if there are pending holds or we're already ready
			if ( wait === true ? --jQuery.readyWait : jQuery.isReady ) {
				return;
			}
	
			// Remember that the DOM is ready
			jQuery.isReady = true;
	
			// If a normal DOM Ready event fired, decrement, and wait if need be
			if ( wait !== true && --jQuery.readyWait > 0 ) {
				return;
			}
	
			// If there are functions bound, to execute
			readyList.resolveWith( document, [ jQuery ] );
		}
	} );
	
	jQuery.ready.then = readyList.then;
	
	// The ready event handler and self cleanup method
	function completed() {
		document.removeEventListener( "DOMContentLoaded", completed );
		window.removeEventListener( "load", completed );
		jQuery.ready();
	}
	
	// Catch cases where $(document).ready() is called
	// after the browser event has already occurred.
	// Support: IE <=9 - 10 only
	// Older IE sometimes signals "interactive" too soon
	if ( document.readyState === "complete" ||
		( document.readyState !== "loading" && !document.documentElement.doScroll ) ) {
	
		// Handle it asynchronously to allow scripts the opportunity to delay ready
		window.setTimeout( jQuery.ready );
	
	} else {
	
		// Use the handy event callback
		document.addEventListener( "DOMContentLoaded", completed );
	
		// A fallback to window.onload, that will always work
		window.addEventListener( "load", completed );
	}
	
	
	
	
	// Multifunctional method to get and set values of a collection
	// The value/s can optionally be executed if it's a function
	var access = function( elems, fn, key, value, chainable, emptyGet, raw ) {
		var i = 0,
			len = elems.length,
			bulk = key == null;
	
		// Sets many values
		if ( jQuery.type( key ) === "object" ) {
			chainable = true;
			for ( i in key ) {
				access( elems, fn, i, key[ i ], true, emptyGet, raw );
			}
	
		// Sets one value
		} else if ( value !== undefined ) {
			chainable = true;
	
			if ( !jQuery.isFunction( value ) ) {
				raw = true;
			}
	
			if ( bulk ) {
	
				// Bulk operations run against the entire set
				if ( raw ) {
					fn.call( elems, value );
					fn = null;
	
				// ...except when executing function values
				} else {
					bulk = fn;
					fn = function( elem, key, value ) {
						return bulk.call( jQuery( elem ), value );
					};
				}
			}
	
			if ( fn ) {
				for ( ; i < len; i++ ) {
					fn(
						elems[ i ], key, raw ?
						value :
						value.call( elems[ i ], i, fn( elems[ i ], key ) )
					);
				}
			}
		}
	
		if ( chainable ) {
			return elems;
		}
	
		// Gets
		if ( bulk ) {
			return fn.call( elems );
		}
	
		return len ? fn( elems[ 0 ], key ) : emptyGet;
	};
	var acceptData = function( owner ) {
	
		// Accepts only:
		//  - Node
		//    - Node.ELEMENT_NODE
		//    - Node.DOCUMENT_NODE
		//  - Object
		//    - Any
		return owner.nodeType === 1 || owner.nodeType === 9 || !( +owner.nodeType );
	};
	
	
	
	
	function Data() {
		this.expando = jQuery.expando + Data.uid++;
	}
	
	Data.uid = 1;
	
	Data.prototype = {
	
		cache: function( owner ) {
	
			// Check if the owner object already has a cache
			var value = owner[ this.expando ];
	
			// If not, create one
			if ( !value ) {
				value = {};
	
				// We can accept data for non-element nodes in modern browsers,
				// but we should not, see #8335.
				// Always return an empty object.
				if ( acceptData( owner ) ) {
	
					// If it is a node unlikely to be stringify-ed or looped over
					// use plain assignment
					if ( owner.nodeType ) {
						owner[ this.expando ] = value;
	
					// Otherwise secure it in a non-enumerable property
					// configurable must be true to allow the property to be
					// deleted when data is removed
					} else {
						Object.defineProperty( owner, this.expando, {
							value: value,
							configurable: true
						} );
					}
				}
			}
	
			return value;
		},
		set: function( owner, data, value ) {
			var prop,
				cache = this.cache( owner );
	
			// Handle: [ owner, key, value ] args
			// Always use camelCase key (gh-2257)
			if ( typeof data === "string" ) {
				cache[ jQuery.camelCase( data ) ] = value;
	
			// Handle: [ owner, { properties } ] args
			} else {
	
				// Copy the properties one-by-one to the cache object
				for ( prop in data ) {
					cache[ jQuery.camelCase( prop ) ] = data[ prop ];
				}
			}
			return cache;
		},
		get: function( owner, key ) {
			return key === undefined ?
				this.cache( owner ) :
	
				// Always use camelCase key (gh-2257)
				owner[ this.expando ] && owner[ this.expando ][ jQuery.camelCase( key ) ];
		},
		access: function( owner, key, value ) {
	
			// In cases where either:
			//
			//   1. No key was specified
			//   2. A string key was specified, but no value provided
			//
			// Take the "read" path and allow the get method to determine
			// which value to return, respectively either:
			//
			//   1. The entire cache object
			//   2. The data stored at the key
			//
			if ( key === undefined ||
					( ( key && typeof key === "string" ) && value === undefined ) ) {
	
				return this.get( owner, key );
			}
	
			// When the key is not a string, or both a key and value
			// are specified, set or extend (existing objects) with either:
			//
			//   1. An object of properties
			//   2. A key and value
			//
			this.set( owner, key, value );
	
			// Since the "set" path can have two possible entry points
			// return the expected data based on which path was taken[*]
			return value !== undefined ? value : key;
		},
		remove: function( owner, key ) {
			var i,
				cache = owner[ this.expando ];
	
			if ( cache === undefined ) {
				return;
			}
	
			if ( key !== undefined ) {
	
				// Support array or space separated string of keys
				if ( Array.isArray( key ) ) {
	
					// If key is an array of keys...
					// We always set camelCase keys, so remove that.
					key = key.map( jQuery.camelCase );
				} else {
					key = jQuery.camelCase( key );
	
					// If a key with the spaces exists, use it.
					// Otherwise, create an array by matching non-whitespace
					key = key in cache ?
						[ key ] :
						( key.match( rnothtmlwhite ) || [] );
				}
	
				i = key.length;
	
				while ( i-- ) {
					delete cache[ key[ i ] ];
				}
			}
	
			// Remove the expando if there's no more data
			if ( key === undefined || jQuery.isEmptyObject( cache ) ) {
	
				// Support: Chrome <=35 - 45
				// Webkit & Blink performance suffers when deleting properties
				// from DOM nodes, so set to undefined instead
				// https://bugs.chromium.org/p/chromium/issues/detail?id=378607 (bug restricted)
				if ( owner.nodeType ) {
					owner[ this.expando ] = undefined;
				} else {
					delete owner[ this.expando ];
				}
			}
		},
		hasData: function( owner ) {
			var cache = owner[ this.expando ];
			return cache !== undefined && !jQuery.isEmptyObject( cache );
		}
	};
	var dataPriv = new Data();
	
	var dataUser = new Data();
	
	
	
	//	Implementation Summary
	//
	//	1. Enforce API surface and semantic compatibility with 1.9.x branch
	//	2. Improve the module's maintainability by reducing the storage
	//		paths to a single mechanism.
	//	3. Use the same single mechanism to support "private" and "user" data.
	//	4. _Never_ expose "private" data to user code (TODO: Drop _data, _removeData)
	//	5. Avoid exposing implementation details on user objects (eg. expando properties)
	//	6. Provide a clear path for implementation upgrade to WeakMap in 2014
	
	var rbrace = /^(?:\{[\w\W]*\}|\[[\w\W]*\])$/,
		rmultiDash = /[A-Z]/g;
	
	function getData( data ) {
		if ( data === "true" ) {
			return true;
		}
	
		if ( data === "false" ) {
			return false;
		}
	
		if ( data === "null" ) {
			return null;
		}
	
		// Only convert to a number if it doesn't change the string
		if ( data === +data + "" ) {
			return +data;
		}
	
		if ( rbrace.test( data ) ) {
			return JSON.parse( data );
		}
	
		return data;
	}
	
	function dataAttr( elem, key, data ) {
		var name;
	
		// If nothing was found internally, try to fetch any
		// data from the HTML5 data-* attribute
		if ( data === undefined && elem.nodeType === 1 ) {
			name = "data-" + key.replace( rmultiDash, "-$&" ).toLowerCase();
			data = elem.getAttribute( name );
	
			if ( typeof data === "string" ) {
				try {
					data = getData( data );
				} catch ( e ) {}
	
				// Make sure we set the data so it isn't changed later
				dataUser.set( elem, key, data );
			} else {
				data = undefined;
			}
		}
		return data;
	}
	
	jQuery.extend( {
		hasData: function( elem ) {
			return dataUser.hasData( elem ) || dataPriv.hasData( elem );
		},
	
		data: function( elem, name, data ) {
			return dataUser.access( elem, name, data );
		},
	
		removeData: function( elem, name ) {
			dataUser.remove( elem, name );
		},
	
		// TODO: Now that all calls to _data and _removeData have been replaced
		// with direct calls to dataPriv methods, these can be deprecated.
		_data: function( elem, name, data ) {
			return dataPriv.access( elem, name, data );
		},
	
		_removeData: function( elem, name ) {
			dataPriv.remove( elem, name );
		}
	} );
	
	jQuery.fn.extend( {
		data: function( key, value ) {
			var i, name, data,
				elem = this[ 0 ],
				attrs = elem && elem.attributes;
	
			// Gets all values
			if ( key === undefined ) {
				if ( this.length ) {
					data = dataUser.get( elem );
	
					if ( elem.nodeType === 1 && !dataPriv.get( elem, "hasDataAttrs" ) ) {
						i = attrs.length;
						while ( i-- ) {
	
							// Support: IE 11 only
							// The attrs elements can be null (#14894)
							if ( attrs[ i ] ) {
								name = attrs[ i ].name;
								if ( name.indexOf( "data-" ) === 0 ) {
									name = jQuery.camelCase( name.slice( 5 ) );
									dataAttr( elem, name, data[ name ] );
								}
							}
						}
						dataPriv.set( elem, "hasDataAttrs", true );
					}
				}
	
				return data;
			}
	
			// Sets multiple values
			if ( typeof key === "object" ) {
				return this.each( function() {
					dataUser.set( this, key );
				} );
			}
	
			return access( this, function( value ) {
				var data;
	
				// The calling jQuery object (element matches) is not empty
				// (and therefore has an element appears at this[ 0 ]) and the
				// `value` parameter was not undefined. An empty jQuery object
				// will result in `undefined` for elem = this[ 0 ] which will
				// throw an exception if an attempt to read a data cache is made.
				if ( elem && value === undefined ) {
	
					// Attempt to get data from the cache
					// The key will always be camelCased in Data
					data = dataUser.get( elem, key );
					if ( data !== undefined ) {
						return data;
					}
	
					// Attempt to "discover" the data in
					// HTML5 custom data-* attrs
					data = dataAttr( elem, key );
					if ( data !== undefined ) {
						return data;
					}
	
					// We tried really hard, but the data doesn't exist.
					return;
				}
	
				// Set the data...
				this.each( function() {
	
					// We always store the camelCased key
					dataUser.set( this, key, value );
				} );
			}, null, value, arguments.length > 1, null, true );
		},
	
		removeData: function( key ) {
			return this.each( function() {
				dataUser.remove( this, key );
			} );
		}
	} );
	
	
	jQuery.extend( {
		queue: function( elem, type, data ) {
			var queue;
	
			if ( elem ) {
				type = ( type || "fx" ) + "queue";
				queue = dataPriv.get( elem, type );
	
				// Speed up dequeue by getting out quickly if this is just a lookup
				if ( data ) {
					if ( !queue || Array.isArray( data ) ) {
						queue = dataPriv.access( elem, type, jQuery.makeArray( data ) );
					} else {
						queue.push( data );
					}
				}
				return queue || [];
			}
		},
	
		dequeue: function( elem, type ) {
			type = type || "fx";
	
			var queue = jQuery.queue( elem, type ),
				startLength = queue.length,
				fn = queue.shift(),
				hooks = jQuery._queueHooks( elem, type ),
				next = function() {
					jQuery.dequeue( elem, type );
				};
	
			// If the fx queue is dequeued, always remove the progress sentinel
			if ( fn === "inprogress" ) {
				fn = queue.shift();
				startLength--;
			}
	
			if ( fn ) {
	
				// Add a progress sentinel to prevent the fx queue from being
				// automatically dequeued
				if ( type === "fx" ) {
					queue.unshift( "inprogress" );
				}
	
				// Clear up the last queue stop function
				delete hooks.stop;
				fn.call( elem, next, hooks );
			}
	
			if ( !startLength && hooks ) {
				hooks.empty.fire();
			}
		},
	
		// Not public - generate a queueHooks object, or return the current one
		_queueHooks: function( elem, type ) {
			var key = type + "queueHooks";
			return dataPriv.get( elem, key ) || dataPriv.access( elem, key, {
				empty: jQuery.Callbacks( "once memory" ).add( function() {
					dataPriv.remove( elem, [ type + "queue", key ] );
				} )
			} );
		}
	} );
	
	jQuery.fn.extend( {
		queue: function( type, data ) {
			var setter = 2;
	
			if ( typeof type !== "string" ) {
				data = type;
				type = "fx";
				setter--;
			}
	
			if ( arguments.length < setter ) {
				return jQuery.queue( this[ 0 ], type );
			}
	
			return data === undefined ?
				this :
				this.each( function() {
					var queue = jQuery.queue( this, type, data );
	
					// Ensure a hooks for this queue
					jQuery._queueHooks( this, type );
	
					if ( type === "fx" && queue[ 0 ] !== "inprogress" ) {
						jQuery.dequeue( this, type );
					}
				} );
		},
		dequeue: function( type ) {
			return this.each( function() {
				jQuery.dequeue( this, type );
			} );
		},
		clearQueue: function( type ) {
			return this.queue( type || "fx", [] );
		},
	
		// Get a promise resolved when queues of a certain type
		// are emptied (fx is the type by default)
		promise: function( type, obj ) {
			var tmp,
				count = 1,
				defer = jQuery.Deferred(),
				elements = this,
				i = this.length,
				resolve = function() {
					if ( !( --count ) ) {
						defer.resolveWith( elements, [ elements ] );
					}
				};
	
			if ( typeof type !== "string" ) {
				obj = type;
				type = undefined;
			}
			type = type || "fx";
	
			while ( i-- ) {
				tmp = dataPriv.get( elements[ i ], type + "queueHooks" );
				if ( tmp && tmp.empty ) {
					count++;
					tmp.empty.add( resolve );
				}
			}
			resolve();
			return defer.promise( obj );
		}
	} );
	var pnum = ( /[+-]?(?:\d*\.|)\d+(?:[eE][+-]?\d+|)/ ).source;
	
	var rcssNum = new RegExp( "^(?:([+-])=|)(" + pnum + ")([a-z%]*)$", "i" );
	
	
	var cssExpand = [ "Top", "Right", "Bottom", "Left" ];
	
	var isHiddenWithinTree = function( elem, el ) {
	
			// isHiddenWithinTree might be called from jQuery#filter function;
			// in that case, element will be second argument
			elem = el || elem;
	
			// Inline style trumps all
			return elem.style.display === "none" ||
				elem.style.display === "" &&
	
				// Otherwise, check computed style
				// Support: Firefox <=43 - 45
				// Disconnected elements can have computed display: none, so first confirm that elem is
				// in the document.
				jQuery.contains( elem.ownerDocument, elem ) &&
	
				jQuery.css( elem, "display" ) === "none";
		};
	
	var swap = function( elem, options, callback, args ) {
		var ret, name,
			old = {};
	
		// Remember the old values, and insert the new ones
		for ( name in options ) {
			old[ name ] = elem.style[ name ];
			elem.style[ name ] = options[ name ];
		}
	
		ret = callback.apply( elem, args || [] );
	
		// Revert the old values
		for ( name in options ) {
			elem.style[ name ] = old[ name ];
		}
	
		return ret;
	};
	
	
	
	
	function adjustCSS( elem, prop, valueParts, tween ) {
		var adjusted,
			scale = 1,
			maxIterations = 20,
			currentValue = tween ?
				function() {
					return tween.cur();
				} :
				function() {
					return jQuery.css( elem, prop, "" );
				},
			initial = currentValue(),
			unit = valueParts && valueParts[ 3 ] || ( jQuery.cssNumber[ prop ] ? "" : "px" ),
	
			// Starting value computation is required for potential unit mismatches
			initialInUnit = ( jQuery.cssNumber[ prop ] || unit !== "px" && +initial ) &&
				rcssNum.exec( jQuery.css( elem, prop ) );
	
		if ( initialInUnit && initialInUnit[ 3 ] !== unit ) {
	
			// Trust units reported by jQuery.css
			unit = unit || initialInUnit[ 3 ];
	
			// Make sure we update the tween properties later on
			valueParts = valueParts || [];
	
			// Iteratively approximate from a nonzero starting point
			initialInUnit = +initial || 1;
	
			do {
	
				// If previous iteration zeroed out, double until we get *something*.
				// Use string for doubling so we don't accidentally see scale as unchanged below
				scale = scale || ".5";
	
				// Adjust and apply
				initialInUnit = initialInUnit / scale;
				jQuery.style( elem, prop, initialInUnit + unit );
	
			// Update scale, tolerating zero or NaN from tween.cur()
			// Break the loop if scale is unchanged or perfect, or if we've just had enough.
			} while (
				scale !== ( scale = currentValue() / initial ) && scale !== 1 && --maxIterations
			);
		}
	
		if ( valueParts ) {
			initialInUnit = +initialInUnit || +initial || 0;
	
			// Apply relative offset (+=/-=) if specified
			adjusted = valueParts[ 1 ] ?
				initialInUnit + ( valueParts[ 1 ] + 1 ) * valueParts[ 2 ] :
				+valueParts[ 2 ];
			if ( tween ) {
				tween.unit = unit;
				tween.start = initialInUnit;
				tween.end = adjusted;
			}
		}
		return adjusted;
	}
	
	
	var defaultDisplayMap = {};
	
	function getDefaultDisplay( elem ) {
		var temp,
			doc = elem.ownerDocument,
			nodeName = elem.nodeName,
			display = defaultDisplayMap[ nodeName ];
	
		if ( display ) {
			return display;
		}
	
		temp = doc.body.appendChild( doc.createElement( nodeName ) );
		display = jQuery.css( temp, "display" );
	
		temp.parentNode.removeChild( temp );
	
		if ( display === "none" ) {
			display = "block";
		}
		defaultDisplayMap[ nodeName ] = display;
	
		return display;
	}
	
	function showHide( elements, show ) {
		var display, elem,
			values = [],
			index = 0,
			length = elements.length;
	
		// Determine new display value for elements that need to change
		for ( ; index < length; index++ ) {
			elem = elements[ index ];
			if ( !elem.style ) {
				continue;
			}
	
			display = elem.style.display;
			if ( show ) {
	
				// Since we force visibility upon cascade-hidden elements, an immediate (and slow)
				// check is required in this first loop unless we have a nonempty display value (either
				// inline or about-to-be-restored)
				if ( display === "none" ) {
					values[ index ] = dataPriv.get( elem, "display" ) || null;
					if ( !values[ index ] ) {
						elem.style.display = "";
					}
				}
				if ( elem.style.display === "" && isHiddenWithinTree( elem ) ) {
					values[ index ] = getDefaultDisplay( elem );
				}
			} else {
				if ( display !== "none" ) {
					values[ index ] = "none";
	
					// Remember what we're overwriting
					dataPriv.set( elem, "display", display );
				}
			}
		}
	
		// Set the display of the elements in a second loop to avoid constant reflow
		for ( index = 0; index < length; index++ ) {
			if ( values[ index ] != null ) {
				elements[ index ].style.display = values[ index ];
			}
		}
	
		return elements;
	}
	
	jQuery.fn.extend( {
		show: function() {
			return showHide( this, true );
		},
		hide: function() {
			return showHide( this );
		},
		toggle: function( state ) {
			if ( typeof state === "boolean" ) {
				return state ? this.show() : this.hide();
			}
	
			return this.each( function() {
				if ( isHiddenWithinTree( this ) ) {
					jQuery( this ).show();
				} else {
					jQuery( this ).hide();
				}
			} );
		}
	} );
	var rcheckableType = ( /^(?:checkbox|radio)$/i );
	
	var rtagName = ( /<([a-z][^\/\0>\x20\t\r\n\f]+)/i );
	
	var rscriptType = ( /^$|\/(?:java|ecma)script/i );
	
	
	
	// We have to close these tags to support XHTML (#13200)
	var wrapMap = {
	
		// Support: IE <=9 only
		option: [ 1, "<select multiple='multiple'>", "</select>" ],
	
		// XHTML parsers do not magically insert elements in the
		// same way that tag soup parsers do. So we cannot shorten
		// this by omitting <tbody> or other required elements.
		thead: [ 1, "<table>", "</table>" ],
		col: [ 2, "<table><colgroup>", "</colgroup></table>" ],
		tr: [ 2, "<table><tbody>", "</tbody></table>" ],
		td: [ 3, "<table><tbody><tr>", "</tr></tbody></table>" ],
	
		_default: [ 0, "", "" ]
	};
	
	// Support: IE <=9 only
	wrapMap.optgroup = wrapMap.option;
	
	wrapMap.tbody = wrapMap.tfoot = wrapMap.colgroup = wrapMap.caption = wrapMap.thead;
	wrapMap.th = wrapMap.td;
	
	
	function getAll( context, tag ) {
	
		// Support: IE <=9 - 11 only
		// Use typeof to avoid zero-argument method invocation on host objects (#15151)
		var ret;
	
		if ( typeof context.getElementsByTagName !== "undefined" ) {
			ret = context.getElementsByTagName( tag || "*" );
	
		} else if ( typeof context.querySelectorAll !== "undefined" ) {
			ret = context.querySelectorAll( tag || "*" );
	
		} else {
			ret = [];
		}
	
		if ( tag === undefined || tag && nodeName( context, tag ) ) {
			return jQuery.merge( [ context ], ret );
		}
	
		return ret;
	}
	
	
	// Mark scripts as having already been evaluated
	function setGlobalEval( elems, refElements ) {
		var i = 0,
			l = elems.length;
	
		for ( ; i < l; i++ ) {
			dataPriv.set(
				elems[ i ],
				"globalEval",
				!refElements || dataPriv.get( refElements[ i ], "globalEval" )
			);
		}
	}
	
	
	var rhtml = /<|&#?\w+;/;
	
	function buildFragment( elems, context, scripts, selection, ignored ) {
		var elem, tmp, tag, wrap, contains, j,
			fragment = context.createDocumentFragment(),
			nodes = [],
			i = 0,
			l = elems.length;
	
		for ( ; i < l; i++ ) {
			elem = elems[ i ];
	
			if ( elem || elem === 0 ) {
	
				// Add nodes directly
				if ( jQuery.type( elem ) === "object" ) {
	
					// Support: Android <=4.0 only, PhantomJS 1 only
					// push.apply(_, arraylike) throws on ancient WebKit
					jQuery.merge( nodes, elem.nodeType ? [ elem ] : elem );
	
				// Convert non-html into a text node
				} else if ( !rhtml.test( elem ) ) {
					nodes.push( context.createTextNode( elem ) );
	
				// Convert html into DOM nodes
				} else {
					tmp = tmp || fragment.appendChild( context.createElement( "div" ) );
	
					// Deserialize a standard representation
					tag = ( rtagName.exec( elem ) || [ "", "" ] )[ 1 ].toLowerCase();
					wrap = wrapMap[ tag ] || wrapMap._default;
					tmp.innerHTML = wrap[ 1 ] + jQuery.htmlPrefilter( elem ) + wrap[ 2 ];
	
					// Descend through wrappers to the right content
					j = wrap[ 0 ];
					while ( j-- ) {
						tmp = tmp.lastChild;
					}
	
					// Support: Android <=4.0 only, PhantomJS 1 only
					// push.apply(_, arraylike) throws on ancient WebKit
					jQuery.merge( nodes, tmp.childNodes );
	
					// Remember the top-level container
					tmp = fragment.firstChild;
	
					// Ensure the created nodes are orphaned (#12392)
					tmp.textContent = "";
				}
			}
		}
	
		// Remove wrapper from fragment
		fragment.textContent = "";
	
		i = 0;
		while ( ( elem = nodes[ i++ ] ) ) {
	
			// Skip elements already in the context collection (trac-4087)
			if ( selection && jQuery.inArray( elem, selection ) > -1 ) {
				if ( ignored ) {
					ignored.push( elem );
				}
				continue;
			}
	
			contains = jQuery.contains( elem.ownerDocument, elem );
	
			// Append to fragment
			tmp = getAll( fragment.appendChild( elem ), "script" );
	
			// Preserve script evaluation history
			if ( contains ) {
				setGlobalEval( tmp );
			}
	
			// Capture executables
			if ( scripts ) {
				j = 0;
				while ( ( elem = tmp[ j++ ] ) ) {
					if ( rscriptType.test( elem.type || "" ) ) {
						scripts.push( elem );
					}
				}
			}
		}
	
		return fragment;
	}
	
	
	( function() {
		var fragment = document.createDocumentFragment(),
			div = fragment.appendChild( document.createElement( "div" ) ),
			input = document.createElement( "input" );
	
		// Support: Android 4.0 - 4.3 only
		// Check state lost if the name is set (#11217)
		// Support: Windows Web Apps (WWA)
		// `name` and `type` must use .setAttribute for WWA (#14901)
		input.setAttribute( "type", "radio" );
		input.setAttribute( "checked", "checked" );
		input.setAttribute( "name", "t" );
	
		div.appendChild( input );
	
		// Support: Android <=4.1 only
		// Older WebKit doesn't clone checked state correctly in fragments
		support.checkClone = div.cloneNode( true ).cloneNode( true ).lastChild.checked;
	
		// Support: IE <=11 only
		// Make sure textarea (and checkbox) defaultValue is properly cloned
		div.innerHTML = "<textarea>x</textarea>";
		support.noCloneChecked = !!div.cloneNode( true ).lastChild.defaultValue;
	} )();
	var documentElement = document.documentElement;
	
	
	
	var
		rkeyEvent = /^key/,
		rmouseEvent = /^(?:mouse|pointer|contextmenu|drag|drop)|click/,
		rtypenamespace = /^([^.]*)(?:\.(.+)|)/;
	
	function returnTrue() {
		return true;
	}
	
	function returnFalse() {
		return false;
	}
	
	// Support: IE <=9 only
	// See #13393 for more info
	function safeActiveElement() {
		try {
			return document.activeElement;
		} catch ( err ) { }
	}
	
	function on( elem, types, selector, data, fn, one ) {
		var origFn, type;
	
		// Types can be a map of types/handlers
		if ( typeof types === "object" ) {
	
			// ( types-Object, selector, data )
			if ( typeof selector !== "string" ) {
	
				// ( types-Object, data )
				data = data || selector;
				selector = undefined;
			}
			for ( type in types ) {
				on( elem, type, selector, data, types[ type ], one );
			}
			return elem;
		}
	
		if ( data == null && fn == null ) {
	
			// ( types, fn )
			fn = selector;
			data = selector = undefined;
		} else if ( fn == null ) {
			if ( typeof selector === "string" ) {
	
				// ( types, selector, fn )
				fn = data;
				data = undefined;
			} else {
	
				// ( types, data, fn )
				fn = data;
				data = selector;
				selector = undefined;
			}
		}
		if ( fn === false ) {
			fn = returnFalse;
		} else if ( !fn ) {
			return elem;
		}
	
		if ( one === 1 ) {
			origFn = fn;
			fn = function( event ) {
	
				// Can use an empty set, since event contains the info
				jQuery().off( event );
				return origFn.apply( this, arguments );
			};
	
			// Use same guid so caller can remove using origFn
			fn.guid = origFn.guid || ( origFn.guid = jQuery.guid++ );
		}
		return elem.each( function() {
			jQuery.event.add( this, types, fn, data, selector );
		} );
	}
	
	/*
	 * Helper functions for managing events -- not part of the public interface.
	 * Props to Dean Edwards' addEvent library for many of the ideas.
	 */
	jQuery.event = {
	
		global: {},
	
		add: function( elem, types, handler, data, selector ) {
	
			var handleObjIn, eventHandle, tmp,
				events, t, handleObj,
				special, handlers, type, namespaces, origType,
				elemData = dataPriv.get( elem );
	
			// Don't attach events to noData or text/comment nodes (but allow plain objects)
			if ( !elemData ) {
				return;
			}
	
			// Caller can pass in an object of custom data in lieu of the handler
			if ( handler.handler ) {
				handleObjIn = handler;
				handler = handleObjIn.handler;
				selector = handleObjIn.selector;
			}
	
			// Ensure that invalid selectors throw exceptions at attach time
			// Evaluate against documentElement in case elem is a non-element node (e.g., document)
			if ( selector ) {
				jQuery.find.matchesSelector( documentElement, selector );
			}
	
			// Make sure that the handler has a unique ID, used to find/remove it later
			if ( !handler.guid ) {
				handler.guid = jQuery.guid++;
			}
	
			// Init the element's event structure and main handler, if this is the first
			if ( !( events = elemData.events ) ) {
				events = elemData.events = {};
			}
			if ( !( eventHandle = elemData.handle ) ) {
				eventHandle = elemData.handle = function( e ) {
	
					// Discard the second event of a jQuery.event.trigger() and
					// when an event is called after a page has unloaded
					return typeof jQuery !== "undefined" && jQuery.event.triggered !== e.type ?
						jQuery.event.dispatch.apply( elem, arguments ) : undefined;
				};
			}
	
			// Handle multiple events separated by a space
			types = ( types || "" ).match( rnothtmlwhite ) || [ "" ];
			t = types.length;
			while ( t-- ) {
				tmp = rtypenamespace.exec( types[ t ] ) || [];
				type = origType = tmp[ 1 ];
				namespaces = ( tmp[ 2 ] || "" ).split( "." ).sort();
	
				// There *must* be a type, no attaching namespace-only handlers
				if ( !type ) {
					continue;
				}
	
				// If event changes its type, use the special event handlers for the changed type
				special = jQuery.event.special[ type ] || {};
	
				// If selector defined, determine special event api type, otherwise given type
				type = ( selector ? special.delegateType : special.bindType ) || type;
	
				// Update special based on newly reset type
				special = jQuery.event.special[ type ] || {};
	
				// handleObj is passed to all event handlers
				handleObj = jQuery.extend( {
					type: type,
					origType: origType,
					data: data,
					handler: handler,
					guid: handler.guid,
					selector: selector,
					needsContext: selector && jQuery.expr.match.needsContext.test( selector ),
					namespace: namespaces.join( "." )
				}, handleObjIn );
	
				// Init the event handler queue if we're the first
				if ( !( handlers = events[ type ] ) ) {
					handlers = events[ type ] = [];
					handlers.delegateCount = 0;
	
					// Only use addEventListener if the special events handler returns false
					if ( !special.setup ||
						special.setup.call( elem, data, namespaces, eventHandle ) === false ) {
	
						if ( elem.addEventListener ) {
							elem.addEventListener( type, eventHandle );
						}
					}
				}
	
				if ( special.add ) {
					special.add.call( elem, handleObj );
	
					if ( !handleObj.handler.guid ) {
						handleObj.handler.guid = handler.guid;
					}
				}
	
				// Add to the element's handler list, delegates in front
				if ( selector ) {
					handlers.splice( handlers.delegateCount++, 0, handleObj );
				} else {
					handlers.push( handleObj );
				}
	
				// Keep track of which events have ever been used, for event optimization
				jQuery.event.global[ type ] = true;
			}
	
		},
	
		// Detach an event or set of events from an element
		remove: function( elem, types, handler, selector, mappedTypes ) {
	
			var j, origCount, tmp,
				events, t, handleObj,
				special, handlers, type, namespaces, origType,
				elemData = dataPriv.hasData( elem ) && dataPriv.get( elem );
	
			if ( !elemData || !( events = elemData.events ) ) {
				return;
			}
	
			// Once for each type.namespace in types; type may be omitted
			types = ( types || "" ).match( rnothtmlwhite ) || [ "" ];
			t = types.length;
			while ( t-- ) {
				tmp = rtypenamespace.exec( types[ t ] ) || [];
				type = origType = tmp[ 1 ];
				namespaces = ( tmp[ 2 ] || "" ).split( "." ).sort();
	
				// Unbind all events (on this namespace, if provided) for the element
				if ( !type ) {
					for ( type in events ) {
						jQuery.event.remove( elem, type + types[ t ], handler, selector, true );
					}
					continue;
				}
	
				special = jQuery.event.special[ type ] || {};
				type = ( selector ? special.delegateType : special.bindType ) || type;
				handlers = events[ type ] || [];
				tmp = tmp[ 2 ] &&
					new RegExp( "(^|\\.)" + namespaces.join( "\\.(?:.*\\.|)" ) + "(\\.|$)" );
	
				// Remove matching events
				origCount = j = handlers.length;
				while ( j-- ) {
					handleObj = handlers[ j ];
	
					if ( ( mappedTypes || origType === handleObj.origType ) &&
						( !handler || handler.guid === handleObj.guid ) &&
						( !tmp || tmp.test( handleObj.namespace ) ) &&
						( !selector || selector === handleObj.selector ||
							selector === "**" && handleObj.selector ) ) {
						handlers.splice( j, 1 );
	
						if ( handleObj.selector ) {
							handlers.delegateCount--;
						}
						if ( special.remove ) {
							special.remove.call( elem, handleObj );
						}
					}
				}
	
				// Remove generic event handler if we removed something and no more handlers exist
				// (avoids potential for endless recursion during removal of special event handlers)
				if ( origCount && !handlers.length ) {
					if ( !special.teardown ||
						special.teardown.call( elem, namespaces, elemData.handle ) === false ) {
	
						jQuery.removeEvent( elem, type, elemData.handle );
					}
	
					delete events[ type ];
				}
			}
	
			// Remove data and the expando if it's no longer used
			if ( jQuery.isEmptyObject( events ) ) {
				dataPriv.remove( elem, "handle events" );
			}
		},
	
		dispatch: function( nativeEvent ) {
	
			// Make a writable jQuery.Event from the native event object
			var event = jQuery.event.fix( nativeEvent );
	
			var i, j, ret, matched, handleObj, handlerQueue,
				args = new Array( arguments.length ),
				handlers = ( dataPriv.get( this, "events" ) || {} )[ event.type ] || [],
				special = jQuery.event.special[ event.type ] || {};
	
			// Use the fix-ed jQuery.Event rather than the (read-only) native event
			args[ 0 ] = event;
	
			for ( i = 1; i < arguments.length; i++ ) {
				args[ i ] = arguments[ i ];
			}
	
			event.delegateTarget = this;
	
			// Call the preDispatch hook for the mapped type, and let it bail if desired
			if ( special.preDispatch && special.preDispatch.call( this, event ) === false ) {
				return;
			}
	
			// Determine handlers
			handlerQueue = jQuery.event.handlers.call( this, event, handlers );
	
			// Run delegates first; they may want to stop propagation beneath us
			i = 0;
			while ( ( matched = handlerQueue[ i++ ] ) && !event.isPropagationStopped() ) {
				event.currentTarget = matched.elem;
	
				j = 0;
				while ( ( handleObj = matched.handlers[ j++ ] ) &&
					!event.isImmediatePropagationStopped() ) {
	
					// Triggered event must either 1) have no namespace, or 2) have namespace(s)
					// a subset or equal to those in the bound event (both can have no namespace).
					if ( !event.rnamespace || event.rnamespace.test( handleObj.namespace ) ) {
	
						event.handleObj = handleObj;
						event.data = handleObj.data;
	
						ret = ( ( jQuery.event.special[ handleObj.origType ] || {} ).handle ||
							handleObj.handler ).apply( matched.elem, args );
	
						if ( ret !== undefined ) {
							if ( ( event.result = ret ) === false ) {
								event.preventDefault();
								event.stopPropagation();
							}
						}
					}
				}
			}
	
			// Call the postDispatch hook for the mapped type
			if ( special.postDispatch ) {
				special.postDispatch.call( this, event );
			}
	
			return event.result;
		},
	
		handlers: function( event, handlers ) {
			var i, handleObj, sel, matchedHandlers, matchedSelectors,
				handlerQueue = [],
				delegateCount = handlers.delegateCount,
				cur = event.target;
	
			// Find delegate handlers
			if ( delegateCount &&
	
				// Support: IE <=9
				// Black-hole SVG <use> instance trees (trac-13180)
				cur.nodeType &&
	
				// Support: Firefox <=42
				// Suppress spec-violating clicks indicating a non-primary pointer button (trac-3861)
				// https://www.w3.org/TR/DOM-Level-3-Events/#event-type-click
				// Support: IE 11 only
				// ...but not arrow key "clicks" of radio inputs, which can have `button` -1 (gh-2343)
				!( event.type === "click" && event.button >= 1 ) ) {
	
				for ( ; cur !== this; cur = cur.parentNode || this ) {
	
					// Don't check non-elements (#13208)
					// Don't process clicks on disabled elements (#6911, #8165, #11382, #11764)
					if ( cur.nodeType === 1 && !( event.type === "click" && cur.disabled === true ) ) {
						matchedHandlers = [];
						matchedSelectors = {};
						for ( i = 0; i < delegateCount; i++ ) {
							handleObj = handlers[ i ];
	
							// Don't conflict with Object.prototype properties (#13203)
							sel = handleObj.selector + " ";
	
							if ( matchedSelectors[ sel ] === undefined ) {
								matchedSelectors[ sel ] = handleObj.needsContext ?
									jQuery( sel, this ).index( cur ) > -1 :
									jQuery.find( sel, this, null, [ cur ] ).length;
							}
							if ( matchedSelectors[ sel ] ) {
								matchedHandlers.push( handleObj );
							}
						}
						if ( matchedHandlers.length ) {
							handlerQueue.push( { elem: cur, handlers: matchedHandlers } );
						}
					}
				}
			}
	
			// Add the remaining (directly-bound) handlers
			cur = this;
			if ( delegateCount < handlers.length ) {
				handlerQueue.push( { elem: cur, handlers: handlers.slice( delegateCount ) } );
			}
	
			return handlerQueue;
		},
	
		addProp: function( name, hook ) {
			Object.defineProperty( jQuery.Event.prototype, name, {
				enumerable: true,
				configurable: true,
	
				get: jQuery.isFunction( hook ) ?
					function() {
						if ( this.originalEvent ) {
								return hook( this.originalEvent );
						}
					} :
					function() {
						if ( this.originalEvent ) {
								return this.originalEvent[ name ];
						}
					},
	
				set: function( value ) {
					Object.defineProperty( this, name, {
						enumerable: true,
						configurable: true,
						writable: true,
						value: value
					} );
				}
			} );
		},
	
		fix: function( originalEvent ) {
			return originalEvent[ jQuery.expando ] ?
				originalEvent :
				new jQuery.Event( originalEvent );
		},
	
		special: {
			load: {
	
				// Prevent triggered image.load events from bubbling to window.load
				noBubble: true
			},
			focus: {
	
				// Fire native event if possible so blur/focus sequence is correct
				trigger: function() {
					if ( this !== safeActiveElement() && this.focus ) {
						this.focus();
						return false;
					}
				},
				delegateType: "focusin"
			},
			blur: {
				trigger: function() {
					if ( this === safeActiveElement() && this.blur ) {
						this.blur();
						return false;
					}
				},
				delegateType: "focusout"
			},
			click: {
	
				// For checkbox, fire native event so checked state will be right
				trigger: function() {
					if ( this.type === "checkbox" && this.click && nodeName( this, "input" ) ) {
						this.click();
						return false;
					}
				},
	
				// For cross-browser consistency, don't fire native .click() on links
				_default: function( event ) {
					return nodeName( event.target, "a" );
				}
			},
	
			beforeunload: {
				postDispatch: function( event ) {
	
					// Support: Firefox 20+
					// Firefox doesn't alert if the returnValue field is not set.
					if ( event.result !== undefined && event.originalEvent ) {
						event.originalEvent.returnValue = event.result;
					}
				}
			}
		}
	};
	
	jQuery.removeEvent = function( elem, type, handle ) {
	
		// This "if" is needed for plain objects
		if ( elem.removeEventListener ) {
			elem.removeEventListener( type, handle );
		}
	};
	
	jQuery.Event = function( src, props ) {
	
		// Allow instantiation without the 'new' keyword
		if ( !( this instanceof jQuery.Event ) ) {
			return new jQuery.Event( src, props );
		}
	
		// Event object
		if ( src && src.type ) {
			this.originalEvent = src;
			this.type = src.type;
	
			// Events bubbling up the document may have been marked as prevented
			// by a handler lower down the tree; reflect the correct value.
			this.isDefaultPrevented = src.defaultPrevented ||
					src.defaultPrevented === undefined &&
	
					// Support: Android <=2.3 only
					src.returnValue === false ?
				returnTrue :
				returnFalse;
	
			// Create target properties
			// Support: Safari <=6 - 7 only
			// Target should not be a text node (#504, #13143)
			this.target = ( src.target && src.target.nodeType === 3 ) ?
				src.target.parentNode :
				src.target;
	
			this.currentTarget = src.currentTarget;
			this.relatedTarget = src.relatedTarget;
	
		// Event type
		} else {
			this.type = src;
		}
	
		// Put explicitly provided properties onto the event object
		if ( props ) {
			jQuery.extend( this, props );
		}
	
		// Create a timestamp if incoming event doesn't have one
		this.timeStamp = src && src.timeStamp || jQuery.now();
	
		// Mark it as fixed
		this[ jQuery.expando ] = true;
	};
	
	// jQuery.Event is based on DOM3 Events as specified by the ECMAScript Language Binding
	// https://www.w3.org/TR/2003/WD-DOM-Level-3-Events-20030331/ecma-script-binding.html
	jQuery.Event.prototype = {
		constructor: jQuery.Event,
		isDefaultPrevented: returnFalse,
		isPropagationStopped: returnFalse,
		isImmediatePropagationStopped: returnFalse,
		isSimulated: false,
	
		preventDefault: function() {
			var e = this.originalEvent;
	
			this.isDefaultPrevented = returnTrue;
	
			if ( e && !this.isSimulated ) {
				e.preventDefault();
			}
		},
		stopPropagation: function() {
			var e = this.originalEvent;
	
			this.isPropagationStopped = returnTrue;
	
			if ( e && !this.isSimulated ) {
				e.stopPropagation();
			}
		},
		stopImmediatePropagation: function() {
			var e = this.originalEvent;
	
			this.isImmediatePropagationStopped = returnTrue;
	
			if ( e && !this.isSimulated ) {
				e.stopImmediatePropagation();
			}
	
			this.stopPropagation();
		}
	};
	
	// Includes all common event props including KeyEvent and MouseEvent specific props
	jQuery.each( {
		altKey: true,
		bubbles: true,
		cancelable: true,
		changedTouches: true,
		ctrlKey: true,
		detail: true,
		eventPhase: true,
		metaKey: true,
		pageX: true,
		pageY: true,
		shiftKey: true,
		view: true,
		"char": true,
		charCode: true,
		key: true,
		keyCode: true,
		button: true,
		buttons: true,
		clientX: true,
		clientY: true,
		offsetX: true,
		offsetY: true,
		pointerId: true,
		pointerType: true,
		screenX: true,
		screenY: true,
		targetTouches: true,
		toElement: true,
		touches: true,
	
		which: function( event ) {
			var button = event.button;
	
			// Add which for key events
			if ( event.which == null && rkeyEvent.test( event.type ) ) {
				return event.charCode != null ? event.charCode : event.keyCode;
			}
	
			// Add which for click: 1 === left; 2 === middle; 3 === right
			if ( !event.which && button !== undefined && rmouseEvent.test( event.type ) ) {
				if ( button & 1 ) {
					return 1;
				}
	
				if ( button & 2 ) {
					return 3;
				}
	
				if ( button & 4 ) {
					return 2;
				}
	
				return 0;
			}
	
			return event.which;
		}
	}, jQuery.event.addProp );
	
	// Create mouseenter/leave events using mouseover/out and event-time checks
	// so that event delegation works in jQuery.
	// Do the same for pointerenter/pointerleave and pointerover/pointerout
	//
	// Support: Safari 7 only
	// Safari sends mouseenter too often; see:
	// https://bugs.chromium.org/p/chromium/issues/detail?id=470258
	// for the description of the bug (it existed in older Chrome versions as well).
	jQuery.each( {
		mouseenter: "mouseover",
		mouseleave: "mouseout",
		pointerenter: "pointerover",
		pointerleave: "pointerout"
	}, function( orig, fix ) {
		jQuery.event.special[ orig ] = {
			delegateType: fix,
			bindType: fix,
	
			handle: function( event ) {
				var ret,
					target = this,
					related = event.relatedTarget,
					handleObj = event.handleObj;
	
				// For mouseenter/leave call the handler if related is outside the target.
				// NB: No relatedTarget if the mouse left/entered the browser window
				if ( !related || ( related !== target && !jQuery.contains( target, related ) ) ) {
					event.type = handleObj.origType;
					ret = handleObj.handler.apply( this, arguments );
					event.type = fix;
				}
				return ret;
			}
		};
	} );
	
	jQuery.fn.extend( {
	
		on: function( types, selector, data, fn ) {
			return on( this, types, selector, data, fn );
		},
		one: function( types, selector, data, fn ) {
			return on( this, types, selector, data, fn, 1 );
		},
		off: function( types, selector, fn ) {
			var handleObj, type;
			if ( types && types.preventDefault && types.handleObj ) {
	
				// ( event )  dispatched jQuery.Event
				handleObj = types.handleObj;
				jQuery( types.delegateTarget ).off(
					handleObj.namespace ?
						handleObj.origType + "." + handleObj.namespace :
						handleObj.origType,
					handleObj.selector,
					handleObj.handler
				);
				return this;
			}
			if ( typeof types === "object" ) {
	
				// ( types-object [, selector] )
				for ( type in types ) {
					this.off( type, selector, types[ type ] );
				}
				return this;
			}
			if ( selector === false || typeof selector === "function" ) {
	
				// ( types [, fn] )
				fn = selector;
				selector = undefined;
			}
			if ( fn === false ) {
				fn = returnFalse;
			}
			return this.each( function() {
				jQuery.event.remove( this, types, fn, selector );
			} );
		}
	} );
	
	
	var
	
		/* eslint-disable max-len */
	
		// See https://github.com/eslint/eslint/issues/3229
		rxhtmlTag = /<(?!area|br|col|embed|hr|img|input|link|meta|param)(([a-z][^\/\0>\x20\t\r\n\f]*)[^>]*)\/>/gi,
	
		/* eslint-enable */
	
		// Support: IE <=10 - 11, Edge 12 - 13
		// In IE/Edge using regex groups here causes severe slowdowns.
		// See https://connect.microsoft.com/IE/feedback/details/1736512/
		rnoInnerhtml = /<script|<style|<link/i,
	
		// checked="checked" or checked
		rchecked = /checked\s*(?:[^=]|=\s*.checked.)/i,
		rscriptTypeMasked = /^true\/(.*)/,
		rcleanScript = /^\s*<!(?:\[CDATA\[|--)|(?:\]\]|--)>\s*$/g;
	
	// Prefer a tbody over its parent table for containing new rows
	function manipulationTarget( elem, content ) {
		if ( nodeName( elem, "table" ) &&
			nodeName( content.nodeType !== 11 ? content : content.firstChild, "tr" ) ) {
	
			return jQuery( ">tbody", elem )[ 0 ] || elem;
		}
	
		return elem;
	}
	
	// Replace/restore the type attribute of script elements for safe DOM manipulation
	function disableScript( elem ) {
		elem.type = ( elem.getAttribute( "type" ) !== null ) + "/" + elem.type;
		return elem;
	}
	function restoreScript( elem ) {
		var match = rscriptTypeMasked.exec( elem.type );
	
		if ( match ) {
			elem.type = match[ 1 ];
		} else {
			elem.removeAttribute( "type" );
		}
	
		return elem;
	}
	
	function cloneCopyEvent( src, dest ) {
		var i, l, type, pdataOld, pdataCur, udataOld, udataCur, events;
	
		if ( dest.nodeType !== 1 ) {
			return;
		}
	
		// 1. Copy private data: events, handlers, etc.
		if ( dataPriv.hasData( src ) ) {
			pdataOld = dataPriv.access( src );
			pdataCur = dataPriv.set( dest, pdataOld );
			events = pdataOld.events;
	
			if ( events ) {
				delete pdataCur.handle;
				pdataCur.events = {};
	
				for ( type in events ) {
					for ( i = 0, l = events[ type ].length; i < l; i++ ) {
						jQuery.event.add( dest, type, events[ type ][ i ] );
					}
				}
			}
		}
	
		// 2. Copy user data
		if ( dataUser.hasData( src ) ) {
			udataOld = dataUser.access( src );
			udataCur = jQuery.extend( {}, udataOld );
	
			dataUser.set( dest, udataCur );
		}
	}
	
	// Fix IE bugs, see support tests
	function fixInput( src, dest ) {
		var nodeName = dest.nodeName.toLowerCase();
	
		// Fails to persist the checked state of a cloned checkbox or radio button.
		if ( nodeName === "input" && rcheckableType.test( src.type ) ) {
			dest.checked = src.checked;
	
		// Fails to return the selected option to the default selected state when cloning options
		} else if ( nodeName === "input" || nodeName === "textarea" ) {
			dest.defaultValue = src.defaultValue;
		}
	}
	
	function domManip( collection, args, callback, ignored ) {
	
		// Flatten any nested arrays
		args = concat.apply( [], args );
	
		var fragment, first, scripts, hasScripts, node, doc,
			i = 0,
			l = collection.length,
			iNoClone = l - 1,
			value = args[ 0 ],
			isFunction = jQuery.isFunction( value );
	
		// We can't cloneNode fragments that contain checked, in WebKit
		if ( isFunction ||
				( l > 1 && typeof value === "string" &&
					!support.checkClone && rchecked.test( value ) ) ) {
			return collection.each( function( index ) {
				var self = collection.eq( index );
				if ( isFunction ) {
					args[ 0 ] = value.call( this, index, self.html() );
				}
				domManip( self, args, callback, ignored );
			} );
		}
	
		if ( l ) {
			fragment = buildFragment( args, collection[ 0 ].ownerDocument, false, collection, ignored );
			first = fragment.firstChild;
	
			if ( fragment.childNodes.length === 1 ) {
				fragment = first;
			}
	
			// Require either new content or an interest in ignored elements to invoke the callback
			if ( first || ignored ) {
				scripts = jQuery.map( getAll( fragment, "script" ), disableScript );
				hasScripts = scripts.length;
	
				// Use the original fragment for the last item
				// instead of the first because it can end up
				// being emptied incorrectly in certain situations (#8070).
				for ( ; i < l; i++ ) {
					node = fragment;
	
					if ( i !== iNoClone ) {
						node = jQuery.clone( node, true, true );
	
						// Keep references to cloned scripts for later restoration
						if ( hasScripts ) {
	
							// Support: Android <=4.0 only, PhantomJS 1 only
							// push.apply(_, arraylike) throws on ancient WebKit
							jQuery.merge( scripts, getAll( node, "script" ) );
						}
					}
	
					callback.call( collection[ i ], node, i );
				}
	
				if ( hasScripts ) {
					doc = scripts[ scripts.length - 1 ].ownerDocument;
	
					// Reenable scripts
					jQuery.map( scripts, restoreScript );
	
					// Evaluate executable scripts on first document insertion
					for ( i = 0; i < hasScripts; i++ ) {
						node = scripts[ i ];
						if ( rscriptType.test( node.type || "" ) &&
							!dataPriv.access( node, "globalEval" ) &&
							jQuery.contains( doc, node ) ) {
	
							if ( node.src ) {
	
								// Optional AJAX dependency, but won't run scripts if not present
								if ( jQuery._evalUrl ) {
									jQuery._evalUrl( node.src );
								}
							} else {
								DOMEval( node.textContent.replace( rcleanScript, "" ), doc );
							}
						}
					}
				}
			}
		}
	
		return collection;
	}
	
	function remove( elem, selector, keepData ) {
		var node,
			nodes = selector ? jQuery.filter( selector, elem ) : elem,
			i = 0;
	
		for ( ; ( node = nodes[ i ] ) != null; i++ ) {
			if ( !keepData && node.nodeType === 1 ) {
				jQuery.cleanData( getAll( node ) );
			}
	
			if ( node.parentNode ) {
				if ( keepData && jQuery.contains( node.ownerDocument, node ) ) {
					setGlobalEval( getAll( node, "script" ) );
				}
				node.parentNode.removeChild( node );
			}
		}
	
		return elem;
	}
	
	jQuery.extend( {
		htmlPrefilter: function( html ) {
			return html.replace( rxhtmlTag, "<$1></$2>" );
		},
	
		clone: function( elem, dataAndEvents, deepDataAndEvents ) {
			var i, l, srcElements, destElements,
				clone = elem.cloneNode( true ),
				inPage = jQuery.contains( elem.ownerDocument, elem );
	
			// Fix IE cloning issues
			if ( !support.noCloneChecked && ( elem.nodeType === 1 || elem.nodeType === 11 ) &&
					!jQuery.isXMLDoc( elem ) ) {
	
				// We eschew Sizzle here for performance reasons: https://jsperf.com/getall-vs-sizzle/2
				destElements = getAll( clone );
				srcElements = getAll( elem );
	
				for ( i = 0, l = srcElements.length; i < l; i++ ) {
					fixInput( srcElements[ i ], destElements[ i ] );
				}
			}
	
			// Copy the events from the original to the clone
			if ( dataAndEvents ) {
				if ( deepDataAndEvents ) {
					srcElements = srcElements || getAll( elem );
					destElements = destElements || getAll( clone );
	
					for ( i = 0, l = srcElements.length; i < l; i++ ) {
						cloneCopyEvent( srcElements[ i ], destElements[ i ] );
					}
				} else {
					cloneCopyEvent( elem, clone );
				}
			}
	
			// Preserve script evaluation history
			destElements = getAll( clone, "script" );
			if ( destElements.length > 0 ) {
				setGlobalEval( destElements, !inPage && getAll( elem, "script" ) );
			}
	
			// Return the cloned set
			return clone;
		},
	
		cleanData: function( elems ) {
			var data, elem, type,
				special = jQuery.event.special,
				i = 0;
	
			for ( ; ( elem = elems[ i ] ) !== undefined; i++ ) {
				if ( acceptData( elem ) ) {
					if ( ( data = elem[ dataPriv.expando ] ) ) {
						if ( data.events ) {
							for ( type in data.events ) {
								if ( special[ type ] ) {
									jQuery.event.remove( elem, type );
	
								// This is a shortcut to avoid jQuery.event.remove's overhead
								} else {
									jQuery.removeEvent( elem, type, data.handle );
								}
							}
						}
	
						// Support: Chrome <=35 - 45+
						// Assign undefined instead of using delete, see Data#remove
						elem[ dataPriv.expando ] = undefined;
					}
					if ( elem[ dataUser.expando ] ) {
	
						// Support: Chrome <=35 - 45+
						// Assign undefined instead of using delete, see Data#remove
						elem[ dataUser.expando ] = undefined;
					}
				}
			}
		}
	} );
	
	jQuery.fn.extend( {
		detach: function( selector ) {
			return remove( this, selector, true );
		},
	
		remove: function( selector ) {
			return remove( this, selector );
		},
	
		text: function( value ) {
			return access( this, function( value ) {
				return value === undefined ?
					jQuery.text( this ) :
					this.empty().each( function() {
						if ( this.nodeType === 1 || this.nodeType === 11 || this.nodeType === 9 ) {
							this.textContent = value;
						}
					} );
			}, null, value, arguments.length );
		},
	
		append: function() {
			return domManip( this, arguments, function( elem ) {
				if ( this.nodeType === 1 || this.nodeType === 11 || this.nodeType === 9 ) {
					var target = manipulationTarget( this, elem );
					target.appendChild( elem );
				}
			} );
		},
	
		prepend: function() {
			return domManip( this, arguments, function( elem ) {
				if ( this.nodeType === 1 || this.nodeType === 11 || this.nodeType === 9 ) {
					var target = manipulationTarget( this, elem );
					target.insertBefore( elem, target.firstChild );
				}
			} );
		},
	
		before: function() {
			return domManip( this, arguments, function( elem ) {
				if ( this.parentNode ) {
					this.parentNode.insertBefore( elem, this );
				}
			} );
		},
	
		after: function() {
			return domManip( this, arguments, function( elem ) {
				if ( this.parentNode ) {
					this.parentNode.insertBefore( elem, this.nextSibling );
				}
			} );
		},
	
		empty: function() {
			var elem,
				i = 0;
	
			for ( ; ( elem = this[ i ] ) != null; i++ ) {
				if ( elem.nodeType === 1 ) {
	
					// Prevent memory leaks
					jQuery.cleanData( getAll( elem, false ) );
	
					// Remove any remaining nodes
					elem.textContent = "";
				}
			}
	
			return this;
		},
	
		clone: function( dataAndEvents, deepDataAndEvents ) {
			dataAndEvents = dataAndEvents == null ? false : dataAndEvents;
			deepDataAndEvents = deepDataAndEvents == null ? dataAndEvents : deepDataAndEvents;
	
			return this.map( function() {
				return jQuery.clone( this, dataAndEvents, deepDataAndEvents );
			} );
		},
	
		html: function( value ) {
			return access( this, function( value ) {
				var elem = this[ 0 ] || {},
					i = 0,
					l = this.length;
	
				if ( value === undefined && elem.nodeType === 1 ) {
					return elem.innerHTML;
				}
	
				// See if we can take a shortcut and just use innerHTML
				if ( typeof value === "string" && !rnoInnerhtml.test( value ) &&
					!wrapMap[ ( rtagName.exec( value ) || [ "", "" ] )[ 1 ].toLowerCase() ] ) {
	
					value = jQuery.htmlPrefilter( value );
	
					try {
						for ( ; i < l; i++ ) {
							elem = this[ i ] || {};
	
							// Remove element nodes and prevent memory leaks
							if ( elem.nodeType === 1 ) {
								jQuery.cleanData( getAll( elem, false ) );
								elem.innerHTML = value;
							}
						}
	
						elem = 0;
	
					// If using innerHTML throws an exception, use the fallback method
					} catch ( e ) {}
				}
	
				if ( elem ) {
					this.empty().append( value );
				}
			}, null, value, arguments.length );
		},
	
		replaceWith: function() {
			var ignored = [];
	
			// Make the changes, replacing each non-ignored context element with the new content
			return domManip( this, arguments, function( elem ) {
				var parent = this.parentNode;
	
				if ( jQuery.inArray( this, ignored ) < 0 ) {
					jQuery.cleanData( getAll( this ) );
					if ( parent ) {
						parent.replaceChild( elem, this );
					}
				}
	
			// Force callback invocation
			}, ignored );
		}
	} );
	
	jQuery.each( {
		appendTo: "append",
		prependTo: "prepend",
		insertBefore: "before",
		insertAfter: "after",
		replaceAll: "replaceWith"
	}, function( name, original ) {
		jQuery.fn[ name ] = function( selector ) {
			var elems,
				ret = [],
				insert = jQuery( selector ),
				last = insert.length - 1,
				i = 0;
	
			for ( ; i <= last; i++ ) {
				elems = i === last ? this : this.clone( true );
				jQuery( insert[ i ] )[ original ]( elems );
	
				// Support: Android <=4.0 only, PhantomJS 1 only
				// .get() because push.apply(_, arraylike) throws on ancient WebKit
				push.apply( ret, elems.get() );
			}
	
			return this.pushStack( ret );
		};
	} );
	var rmargin = ( /^margin/ );
	
	var rnumnonpx = new RegExp( "^(" + pnum + ")(?!px)[a-z%]+$", "i" );
	
	var getStyles = function( elem ) {
	
			// Support: IE <=11 only, Firefox <=30 (#15098, #14150)
			// IE throws on elements created in popups
			// FF meanwhile throws on frame elements through "defaultView.getComputedStyle"
			var view = elem.ownerDocument.defaultView;
	
			if ( !view || !view.opener ) {
				view = window;
			}
	
			return view.getComputedStyle( elem );
		};
	
	
	
	( function() {
	
		// Executing both pixelPosition & boxSizingReliable tests require only one layout
		// so they're executed at the same time to save the second computation.
		function computeStyleTests() {
	
			// This is a singleton, we need to execute it only once
			if ( !div ) {
				return;
			}
	
			div.style.cssText =
				"box-sizing:border-box;" +
				"position:relative;display:block;" +
				"margin:auto;border:1px;padding:1px;" +
				"top:1%;width:50%";
			div.innerHTML = "";
			documentElement.appendChild( container );
	
			var divStyle = window.getComputedStyle( div );
			pixelPositionVal = divStyle.top !== "1%";
	
			// Support: Android 4.0 - 4.3 only, Firefox <=3 - 44
			reliableMarginLeftVal = divStyle.marginLeft === "2px";
			boxSizingReliableVal = divStyle.width === "4px";
	
			// Support: Android 4.0 - 4.3 only
			// Some styles come back with percentage values, even though they shouldn't
			div.style.marginRight = "50%";
			pixelMarginRightVal = divStyle.marginRight === "4px";
	
			documentElement.removeChild( container );
	
			// Nullify the div so it wouldn't be stored in the memory and
			// it will also be a sign that checks already performed
			div = null;
		}
	
		var pixelPositionVal, boxSizingReliableVal, pixelMarginRightVal, reliableMarginLeftVal,
			container = document.createElement( "div" ),
			div = document.createElement( "div" );
	
		// Finish early in limited (non-browser) environments
		if ( !div.style ) {
			return;
		}
	
		// Support: IE <=9 - 11 only
		// Style of cloned element affects source element cloned (#8908)
		div.style.backgroundClip = "content-box";
		div.cloneNode( true ).style.backgroundClip = "";
		support.clearCloneStyle = div.style.backgroundClip === "content-box";
	
		container.style.cssText = "border:0;width:8px;height:0;top:0;left:-9999px;" +
			"padding:0;margin-top:1px;position:absolute";
		container.appendChild( div );
	
		jQuery.extend( support, {
			pixelPosition: function() {
				computeStyleTests();
				return pixelPositionVal;
			},
			boxSizingReliable: function() {
				computeStyleTests();
				return boxSizingReliableVal;
			},
			pixelMarginRight: function() {
				computeStyleTests();
				return pixelMarginRightVal;
			},
			reliableMarginLeft: function() {
				computeStyleTests();
				return reliableMarginLeftVal;
			}
		} );
	} )();
	
	
	function curCSS( elem, name, computed ) {
		var width, minWidth, maxWidth, ret,
	
			// Support: Firefox 51+
			// Retrieving style before computed somehow
			// fixes an issue with getting wrong values
			// on detached elements
			style = elem.style;
	
		computed = computed || getStyles( elem );
	
		// getPropertyValue is needed for:
		//   .css('filter') (IE 9 only, #12537)
		//   .css('--customProperty) (#3144)
		if ( computed ) {
			ret = computed.getPropertyValue( name ) || computed[ name ];
	
			if ( ret === "" && !jQuery.contains( elem.ownerDocument, elem ) ) {
				ret = jQuery.style( elem, name );
			}
	
			// A tribute to the "awesome hack by Dean Edwards"
			// Android Browser returns percentage for some values,
			// but width seems to be reliably pixels.
			// This is against the CSSOM draft spec:
			// https://drafts.csswg.org/cssom/#resolved-values
			if ( !support.pixelMarginRight() && rnumnonpx.test( ret ) && rmargin.test( name ) ) {
	
				// Remember the original values
				width = style.width;
				minWidth = style.minWidth;
				maxWidth = style.maxWidth;
	
				// Put in the new values to get a computed value out
				style.minWidth = style.maxWidth = style.width = ret;
				ret = computed.width;
	
				// Revert the changed values
				style.width = width;
				style.minWidth = minWidth;
				style.maxWidth = maxWidth;
			}
		}
	
		return ret !== undefined ?
	
			// Support: IE <=9 - 11 only
			// IE returns zIndex value as an integer.
			ret + "" :
			ret;
	}
	
	
	function addGetHookIf( conditionFn, hookFn ) {
	
		// Define the hook, we'll check on the first run if it's really needed.
		return {
			get: function() {
				if ( conditionFn() ) {
	
					// Hook not needed (or it's not possible to use it due
					// to missing dependency), remove it.
					delete this.get;
					return;
				}
	
				// Hook needed; redefine it so that the support test is not executed again.
				return ( this.get = hookFn ).apply( this, arguments );
			}
		};
	}
	
	
	var
	
		// Swappable if display is none or starts with table
		// except "table", "table-cell", or "table-caption"
		// See here for display values: https://developer.mozilla.org/en-US/docs/CSS/display
		rdisplayswap = /^(none|table(?!-c[ea]).+)/,
		rcustomProp = /^--/,
		cssShow = { position: "absolute", visibility: "hidden", display: "block" },
		cssNormalTransform = {
			letterSpacing: "0",
			fontWeight: "400"
		},
	
		cssPrefixes = [ "Webkit", "Moz", "ms" ],
		emptyStyle = document.createElement( "div" ).style;
	
	// Return a css property mapped to a potentially vendor prefixed property
	function vendorPropName( name ) {
	
		// Shortcut for names that are not vendor prefixed
		if ( name in emptyStyle ) {
			return name;
		}
	
		// Check for vendor prefixed names
		var capName = name[ 0 ].toUpperCase() + name.slice( 1 ),
			i = cssPrefixes.length;
	
		while ( i-- ) {
			name = cssPrefixes[ i ] + capName;
			if ( name in emptyStyle ) {
				return name;
			}
		}
	}
	
	// Return a property mapped along what jQuery.cssProps suggests or to
	// a vendor prefixed property.
	function finalPropName( name ) {
		var ret = jQuery.cssProps[ name ];
		if ( !ret ) {
			ret = jQuery.cssProps[ name ] = vendorPropName( name ) || name;
		}
		return ret;
	}
	
	function setPositiveNumber( elem, value, subtract ) {
	
		// Any relative (+/-) values have already been
		// normalized at this point
		var matches = rcssNum.exec( value );
		return matches ?
	
			// Guard against undefined "subtract", e.g., when used as in cssHooks
			Math.max( 0, matches[ 2 ] - ( subtract || 0 ) ) + ( matches[ 3 ] || "px" ) :
			value;
	}
	
	function augmentWidthOrHeight( elem, name, extra, isBorderBox, styles ) {
		var i,
			val = 0;
	
		// If we already have the right measurement, avoid augmentation
		if ( extra === ( isBorderBox ? "border" : "content" ) ) {
			i = 4;
	
		// Otherwise initialize for horizontal or vertical properties
		} else {
			i = name === "width" ? 1 : 0;
		}
	
		for ( ; i < 4; i += 2 ) {
	
			// Both box models exclude margin, so add it if we want it
			if ( extra === "margin" ) {
				val += jQuery.css( elem, extra + cssExpand[ i ], true, styles );
			}
	
			if ( isBorderBox ) {
	
				// border-box includes padding, so remove it if we want content
				if ( extra === "content" ) {
					val -= jQuery.css( elem, "padding" + cssExpand[ i ], true, styles );
				}
	
				// At this point, extra isn't border nor margin, so remove border
				if ( extra !== "margin" ) {
					val -= jQuery.css( elem, "border" + cssExpand[ i ] + "Width", true, styles );
				}
			} else {
	
				// At this point, extra isn't content, so add padding
				val += jQuery.css( elem, "padding" + cssExpand[ i ], true, styles );
	
				// At this point, extra isn't content nor padding, so add border
				if ( extra !== "padding" ) {
					val += jQuery.css( elem, "border" + cssExpand[ i ] + "Width", true, styles );
				}
			}
		}
	
		return val;
	}
	
	function getWidthOrHeight( elem, name, extra ) {
	
		// Start with computed style
		var valueIsBorderBox,
			styles = getStyles( elem ),
			val = curCSS( elem, name, styles ),
			isBorderBox = jQuery.css( elem, "boxSizing", false, styles ) === "border-box";
	
		// Computed unit is not pixels. Stop here and return.
		if ( rnumnonpx.test( val ) ) {
			return val;
		}
	
		// Check for style in case a browser which returns unreliable values
		// for getComputedStyle silently falls back to the reliable elem.style
		valueIsBorderBox = isBorderBox &&
			( support.boxSizingReliable() || val === elem.style[ name ] );
	
		// Fall back to offsetWidth/Height when value is "auto"
		// This happens for inline elements with no explicit setting (gh-3571)
		if ( val === "auto" ) {
			val = elem[ "offset" + name[ 0 ].toUpperCase() + name.slice( 1 ) ];
		}
	
		// Normalize "", auto, and prepare for extra
		val = parseFloat( val ) || 0;
	
		// Use the active box-sizing model to add/subtract irrelevant styles
		return ( val +
			augmentWidthOrHeight(
				elem,
				name,
				extra || ( isBorderBox ? "border" : "content" ),
				valueIsBorderBox,
				styles
			)
		) + "px";
	}
	
	jQuery.extend( {
	
		// Add in style property hooks for overriding the default
		// behavior of getting and setting a style property
		cssHooks: {
			opacity: {
				get: function( elem, computed ) {
					if ( computed ) {
	
						// We should always get a number back from opacity
						var ret = curCSS( elem, "opacity" );
						return ret === "" ? "1" : ret;
					}
				}
			}
		},
	
		// Don't automatically add "px" to these possibly-unitless properties
		cssNumber: {
			"animationIterationCount": true,
			"columnCount": true,
			"fillOpacity": true,
			"flexGrow": true,
			"flexShrink": true,
			"fontWeight": true,
			"lineHeight": true,
			"opacity": true,
			"order": true,
			"orphans": true,
			"widows": true,
			"zIndex": true,
			"zoom": true
		},
	
		// Add in properties whose names you wish to fix before
		// setting or getting the value
		cssProps: {
			"float": "cssFloat"
		},
	
		// Get and set the style property on a DOM Node
		style: function( elem, name, value, extra ) {
	
			// Don't set styles on text and comment nodes
			if ( !elem || elem.nodeType === 3 || elem.nodeType === 8 || !elem.style ) {
				return;
			}
	
			// Make sure that we're working with the right name
			var ret, type, hooks,
				origName = jQuery.camelCase( name ),
				isCustomProp = rcustomProp.test( name ),
				style = elem.style;
	
			// Make sure that we're working with the right name. We don't
			// want to query the value if it is a CSS custom property
			// since they are user-defined.
			if ( !isCustomProp ) {
				name = finalPropName( origName );
			}
	
			// Gets hook for the prefixed version, then unprefixed version
			hooks = jQuery.cssHooks[ name ] || jQuery.cssHooks[ origName ];
	
			// Check if we're setting a value
			if ( value !== undefined ) {
				type = typeof value;
	
				// Convert "+=" or "-=" to relative numbers (#7345)
				if ( type === "string" && ( ret = rcssNum.exec( value ) ) && ret[ 1 ] ) {
					value = adjustCSS( elem, name, ret );
	
					// Fixes bug #9237
					type = "number";
				}
	
				// Make sure that null and NaN values aren't set (#7116)
				if ( value == null || value !== value ) {
					return;
				}
	
				// If a number was passed in, add the unit (except for certain CSS properties)
				if ( type === "number" ) {
					value += ret && ret[ 3 ] || ( jQuery.cssNumber[ origName ] ? "" : "px" );
				}
	
				// background-* props affect original clone's values
				if ( !support.clearCloneStyle && value === "" && name.indexOf( "background" ) === 0 ) {
					style[ name ] = "inherit";
				}
	
				// If a hook was provided, use that value, otherwise just set the specified value
				if ( !hooks || !( "set" in hooks ) ||
					( value = hooks.set( elem, value, extra ) ) !== undefined ) {
	
					if ( isCustomProp ) {
						style.setProperty( name, value );
					} else {
						style[ name ] = value;
					}
				}
	
			} else {
	
				// If a hook was provided get the non-computed value from there
				if ( hooks && "get" in hooks &&
					( ret = hooks.get( elem, false, extra ) ) !== undefined ) {
	
					return ret;
				}
	
				// Otherwise just get the value from the style object
				return style[ name ];
			}
		},
	
		css: function( elem, name, extra, styles ) {
			var val, num, hooks,
				origName = jQuery.camelCase( name ),
				isCustomProp = rcustomProp.test( name );
	
			// Make sure that we're working with the right name. We don't
			// want to modify the value if it is a CSS custom property
			// since they are user-defined.
			if ( !isCustomProp ) {
				name = finalPropName( origName );
			}
	
			// Try prefixed name followed by the unprefixed name
			hooks = jQuery.cssHooks[ name ] || jQuery.cssHooks[ origName ];
	
			// If a hook was provided get the computed value from there
			if ( hooks && "get" in hooks ) {
				val = hooks.get( elem, true, extra );
			}
	
			// Otherwise, if a way to get the computed value exists, use that
			if ( val === undefined ) {
				val = curCSS( elem, name, styles );
			}
	
			// Convert "normal" to computed value
			if ( val === "normal" && name in cssNormalTransform ) {
				val = cssNormalTransform[ name ];
			}
	
			// Make numeric if forced or a qualifier was provided and val looks numeric
			if ( extra === "" || extra ) {
				num = parseFloat( val );
				return extra === true || isFinite( num ) ? num || 0 : val;
			}
	
			return val;
		}
	} );
	
	jQuery.each( [ "height", "width" ], function( i, name ) {
		jQuery.cssHooks[ name ] = {
			get: function( elem, computed, extra ) {
				if ( computed ) {
	
					// Certain elements can have dimension info if we invisibly show them
					// but it must have a current display style that would benefit
					return rdisplayswap.test( jQuery.css( elem, "display" ) ) &&
	
						// Support: Safari 8+
						// Table columns in Safari have non-zero offsetWidth & zero
						// getBoundingClientRect().width unless display is changed.
						// Support: IE <=11 only
						// Running getBoundingClientRect on a disconnected node
						// in IE throws an error.
						( !elem.getClientRects().length || !elem.getBoundingClientRect().width ) ?
							swap( elem, cssShow, function() {
								return getWidthOrHeight( elem, name, extra );
							} ) :
							getWidthOrHeight( elem, name, extra );
				}
			},
	
			set: function( elem, value, extra ) {
				var matches,
					styles = extra && getStyles( elem ),
					subtract = extra && augmentWidthOrHeight(
						elem,
						name,
						extra,
						jQuery.css( elem, "boxSizing", false, styles ) === "border-box",
						styles
					);
	
				// Convert to pixels if value adjustment is needed
				if ( subtract && ( matches = rcssNum.exec( value ) ) &&
					( matches[ 3 ] || "px" ) !== "px" ) {
	
					elem.style[ name ] = value;
					value = jQuery.css( elem, name );
				}
	
				return setPositiveNumber( elem, value, subtract );
			}
		};
	} );
	
	jQuery.cssHooks.marginLeft = addGetHookIf( support.reliableMarginLeft,
		function( elem, computed ) {
			if ( computed ) {
				return ( parseFloat( curCSS( elem, "marginLeft" ) ) ||
					elem.getBoundingClientRect().left -
						swap( elem, { marginLeft: 0 }, function() {
							return elem.getBoundingClientRect().left;
						} )
					) + "px";
			}
		}
	);
	
	// These hooks are used by animate to expand properties
	jQuery.each( {
		margin: "",
		padding: "",
		border: "Width"
	}, function( prefix, suffix ) {
		jQuery.cssHooks[ prefix + suffix ] = {
			expand: function( value ) {
				var i = 0,
					expanded = {},
	
					// Assumes a single number if not a string
					parts = typeof value === "string" ? value.split( " " ) : [ value ];
	
				for ( ; i < 4; i++ ) {
					expanded[ prefix + cssExpand[ i ] + suffix ] =
						parts[ i ] || parts[ i - 2 ] || parts[ 0 ];
				}
	
				return expanded;
			}
		};
	
		if ( !rmargin.test( prefix ) ) {
			jQuery.cssHooks[ prefix + suffix ].set = setPositiveNumber;
		}
	} );
	
	jQuery.fn.extend( {
		css: function( name, value ) {
			return access( this, function( elem, name, value ) {
				var styles, len,
					map = {},
					i = 0;
	
				if ( Array.isArray( name ) ) {
					styles = getStyles( elem );
					len = name.length;
	
					for ( ; i < len; i++ ) {
						map[ name[ i ] ] = jQuery.css( elem, name[ i ], false, styles );
					}
	
					return map;
				}
	
				return value !== undefined ?
					jQuery.style( elem, name, value ) :
					jQuery.css( elem, name );
			}, name, value, arguments.length > 1 );
		}
	} );
	
	
	function Tween( elem, options, prop, end, easing ) {
		return new Tween.prototype.init( elem, options, prop, end, easing );
	}
	jQuery.Tween = Tween;
	
	Tween.prototype = {
		constructor: Tween,
		init: function( elem, options, prop, end, easing, unit ) {
			this.elem = elem;
			this.prop = prop;
			this.easing = easing || jQuery.easing._default;
			this.options = options;
			this.start = this.now = this.cur();
			this.end = end;
			this.unit = unit || ( jQuery.cssNumber[ prop ] ? "" : "px" );
		},
		cur: function() {
			var hooks = Tween.propHooks[ this.prop ];
	
			return hooks && hooks.get ?
				hooks.get( this ) :
				Tween.propHooks._default.get( this );
		},
		run: function( percent ) {
			var eased,
				hooks = Tween.propHooks[ this.prop ];
	
			if ( this.options.duration ) {
				this.pos = eased = jQuery.easing[ this.easing ](
					percent, this.options.duration * percent, 0, 1, this.options.duration
				);
			} else {
				this.pos = eased = percent;
			}
			this.now = ( this.end - this.start ) * eased + this.start;
	
			if ( this.options.step ) {
				this.options.step.call( this.elem, this.now, this );
			}
	
			if ( hooks && hooks.set ) {
				hooks.set( this );
			} else {
				Tween.propHooks._default.set( this );
			}
			return this;
		}
	};
	
	Tween.prototype.init.prototype = Tween.prototype;
	
	Tween.propHooks = {
		_default: {
			get: function( tween ) {
				var result;
	
				// Use a property on the element directly when it is not a DOM element,
				// or when there is no matching style property that exists.
				if ( tween.elem.nodeType !== 1 ||
					tween.elem[ tween.prop ] != null && tween.elem.style[ tween.prop ] == null ) {
					return tween.elem[ tween.prop ];
				}
	
				// Passing an empty string as a 3rd parameter to .css will automatically
				// attempt a parseFloat and fallback to a string if the parse fails.
				// Simple values such as "10px" are parsed to Float;
				// complex values such as "rotate(1rad)" are returned as-is.
				result = jQuery.css( tween.elem, tween.prop, "" );
	
				// Empty strings, null, undefined and "auto" are converted to 0.
				return !result || result === "auto" ? 0 : result;
			},
			set: function( tween ) {
	
				// Use step hook for back compat.
				// Use cssHook if its there.
				// Use .style if available and use plain properties where available.
				if ( jQuery.fx.step[ tween.prop ] ) {
					jQuery.fx.step[ tween.prop ]( tween );
				} else if ( tween.elem.nodeType === 1 &&
					( tween.elem.style[ jQuery.cssProps[ tween.prop ] ] != null ||
						jQuery.cssHooks[ tween.prop ] ) ) {
					jQuery.style( tween.elem, tween.prop, tween.now + tween.unit );
				} else {
					tween.elem[ tween.prop ] = tween.now;
				}
			}
		}
	};
	
	// Support: IE <=9 only
	// Panic based approach to setting things on disconnected nodes
	Tween.propHooks.scrollTop = Tween.propHooks.scrollLeft = {
		set: function( tween ) {
			if ( tween.elem.nodeType && tween.elem.parentNode ) {
				tween.elem[ tween.prop ] = tween.now;
			}
		}
	};
	
	jQuery.easing = {
		linear: function( p ) {
			return p;
		},
		swing: function( p ) {
			return 0.5 - Math.cos( p * Math.PI ) / 2;
		},
		_default: "swing"
	};
	
	jQuery.fx = Tween.prototype.init;
	
	// Back compat <1.8 extension point
	jQuery.fx.step = {};
	
	
	
	
	var
		fxNow, inProgress,
		rfxtypes = /^(?:toggle|show|hide)$/,
		rrun = /queueHooks$/;
	
	function schedule() {
		if ( inProgress ) {
			if ( document.hidden === false && window.requestAnimationFrame ) {
				window.requestAnimationFrame( schedule );
			} else {
				window.setTimeout( schedule, jQuery.fx.interval );
			}
	
			jQuery.fx.tick();
		}
	}
	
	// Animations created synchronously will run synchronously
	function createFxNow() {
		window.setTimeout( function() {
			fxNow = undefined;
		} );
		return ( fxNow = jQuery.now() );
	}
	
	// Generate parameters to create a standard animation
	function genFx( type, includeWidth ) {
		var which,
			i = 0,
			attrs = { height: type };
	
		// If we include width, step value is 1 to do all cssExpand values,
		// otherwise step value is 2 to skip over Left and Right
		includeWidth = includeWidth ? 1 : 0;
		for ( ; i < 4; i += 2 - includeWidth ) {
			which = cssExpand[ i ];
			attrs[ "margin" + which ] = attrs[ "padding" + which ] = type;
		}
	
		if ( includeWidth ) {
			attrs.opacity = attrs.width = type;
		}
	
		return attrs;
	}
	
	function createTween( value, prop, animation ) {
		var tween,
			collection = ( Animation.tweeners[ prop ] || [] ).concat( Animation.tweeners[ "*" ] ),
			index = 0,
			length = collection.length;
		for ( ; index < length; index++ ) {
			if ( ( tween = collection[ index ].call( animation, prop, value ) ) ) {
	
				// We're done with this property
				return tween;
			}
		}
	}
	
	function defaultPrefilter( elem, props, opts ) {
		var prop, value, toggle, hooks, oldfire, propTween, restoreDisplay, display,
			isBox = "width" in props || "height" in props,
			anim = this,
			orig = {},
			style = elem.style,
			hidden = elem.nodeType && isHiddenWithinTree( elem ),
			dataShow = dataPriv.get( elem, "fxshow" );
	
		// Queue-skipping animations hijack the fx hooks
		if ( !opts.queue ) {
			hooks = jQuery._queueHooks( elem, "fx" );
			if ( hooks.unqueued == null ) {
				hooks.unqueued = 0;
				oldfire = hooks.empty.fire;
				hooks.empty.fire = function() {
					if ( !hooks.unqueued ) {
						oldfire();
					}
				};
			}
			hooks.unqueued++;
	
			anim.always( function() {
	
				// Ensure the complete handler is called before this completes
				anim.always( function() {
					hooks.unqueued--;
					if ( !jQuery.queue( elem, "fx" ).length ) {
						hooks.empty.fire();
					}
				} );
			} );
		}
	
		// Detect show/hide animations
		for ( prop in props ) {
			value = props[ prop ];
			if ( rfxtypes.test( value ) ) {
				delete props[ prop ];
				toggle = toggle || value === "toggle";
				if ( value === ( hidden ? "hide" : "show" ) ) {
	
					// Pretend to be hidden if this is a "show" and
					// there is still data from a stopped show/hide
					if ( value === "show" && dataShow && dataShow[ prop ] !== undefined ) {
						hidden = true;
	
					// Ignore all other no-op show/hide data
					} else {
						continue;
					}
				}
				orig[ prop ] = dataShow && dataShow[ prop ] || jQuery.style( elem, prop );
			}
		}
	
		// Bail out if this is a no-op like .hide().hide()
		propTween = !jQuery.isEmptyObject( props );
		if ( !propTween && jQuery.isEmptyObject( orig ) ) {
			return;
		}
	
		// Restrict "overflow" and "display" styles during box animations
		if ( isBox && elem.nodeType === 1 ) {
	
			// Support: IE <=9 - 11, Edge 12 - 13
			// Record all 3 overflow attributes because IE does not infer the shorthand
			// from identically-valued overflowX and overflowY
			opts.overflow = [ style.overflow, style.overflowX, style.overflowY ];
	
			// Identify a display type, preferring old show/hide data over the CSS cascade
			restoreDisplay = dataShow && dataShow.display;
			if ( restoreDisplay == null ) {
				restoreDisplay = dataPriv.get( elem, "display" );
			}
			display = jQuery.css( elem, "display" );
			if ( display === "none" ) {
				if ( restoreDisplay ) {
					display = restoreDisplay;
				} else {
	
					// Get nonempty value(s) by temporarily forcing visibility
					showHide( [ elem ], true );
					restoreDisplay = elem.style.display || restoreDisplay;
					display = jQuery.css( elem, "display" );
					showHide( [ elem ] );
				}
			}
	
			// Animate inline elements as inline-block
			if ( display === "inline" || display === "inline-block" && restoreDisplay != null ) {
				if ( jQuery.css( elem, "float" ) === "none" ) {
	
					// Restore the original display value at the end of pure show/hide animations
					if ( !propTween ) {
						anim.done( function() {
							style.display = restoreDisplay;
						} );
						if ( restoreDisplay == null ) {
							display = style.display;
							restoreDisplay = display === "none" ? "" : display;
						}
					}
					style.display = "inline-block";
				}
			}
		}
	
		if ( opts.overflow ) {
			style.overflow = "hidden";
			anim.always( function() {
				style.overflow = opts.overflow[ 0 ];
				style.overflowX = opts.overflow[ 1 ];
				style.overflowY = opts.overflow[ 2 ];
			} );
		}
	
		// Implement show/hide animations
		propTween = false;
		for ( prop in orig ) {
	
			// General show/hide setup for this element animation
			if ( !propTween ) {
				if ( dataShow ) {
					if ( "hidden" in dataShow ) {
						hidden = dataShow.hidden;
					}
				} else {
					dataShow = dataPriv.access( elem, "fxshow", { display: restoreDisplay } );
				}
	
				// Store hidden/visible for toggle so `.stop().toggle()` "reverses"
				if ( toggle ) {
					dataShow.hidden = !hidden;
				}
	
				// Show elements before animating them
				if ( hidden ) {
					showHide( [ elem ], true );
				}
	
				/* eslint-disable no-loop-func */
	
				anim.done( function() {
	
				/* eslint-enable no-loop-func */
	
					// The final step of a "hide" animation is actually hiding the element
					if ( !hidden ) {
						showHide( [ elem ] );
					}
					dataPriv.remove( elem, "fxshow" );
					for ( prop in orig ) {
						jQuery.style( elem, prop, orig[ prop ] );
					}
				} );
			}
	
			// Per-property setup
			propTween = createTween( hidden ? dataShow[ prop ] : 0, prop, anim );
			if ( !( prop in dataShow ) ) {
				dataShow[ prop ] = propTween.start;
				if ( hidden ) {
					propTween.end = propTween.start;
					propTween.start = 0;
				}
			}
		}
	}
	
	function propFilter( props, specialEasing ) {
		var index, name, easing, value, hooks;
	
		// camelCase, specialEasing and expand cssHook pass
		for ( index in props ) {
			name = jQuery.camelCase( index );
			easing = specialEasing[ name ];
			value = props[ index ];
			if ( Array.isArray( value ) ) {
				easing = value[ 1 ];
				value = props[ index ] = value[ 0 ];
			}
	
			if ( index !== name ) {
				props[ name ] = value;
				delete props[ index ];
			}
	
			hooks = jQuery.cssHooks[ name ];
			if ( hooks && "expand" in hooks ) {
				value = hooks.expand( value );
				delete props[ name ];
	
				// Not quite $.extend, this won't overwrite existing keys.
				// Reusing 'index' because we have the correct "name"
				for ( index in value ) {
					if ( !( index in props ) ) {
						props[ index ] = value[ index ];
						specialEasing[ index ] = easing;
					}
				}
			} else {
				specialEasing[ name ] = easing;
			}
		}
	}
	
	function Animation( elem, properties, options ) {
		var result,
			stopped,
			index = 0,
			length = Animation.prefilters.length,
			deferred = jQuery.Deferred().always( function() {
	
				// Don't match elem in the :animated selector
				delete tick.elem;
			} ),
			tick = function() {
				if ( stopped ) {
					return false;
				}
				var currentTime = fxNow || createFxNow(),
					remaining = Math.max( 0, animation.startTime + animation.duration - currentTime ),
	
					// Support: Android 2.3 only
					// Archaic crash bug won't allow us to use `1 - ( 0.5 || 0 )` (#12497)
					temp = remaining / animation.duration || 0,
					percent = 1 - temp,
					index = 0,
					length = animation.tweens.length;
	
				for ( ; index < length; index++ ) {
					animation.tweens[ index ].run( percent );
				}
	
				deferred.notifyWith( elem, [ animation, percent, remaining ] );
	
				// If there's more to do, yield
				if ( percent < 1 && length ) {
					return remaining;
				}
	
				// If this was an empty animation, synthesize a final progress notification
				if ( !length ) {
					deferred.notifyWith( elem, [ animation, 1, 0 ] );
				}
	
				// Resolve the animation and report its conclusion
				deferred.resolveWith( elem, [ animation ] );
				return false;
			},
			animation = deferred.promise( {
				elem: elem,
				props: jQuery.extend( {}, properties ),
				opts: jQuery.extend( true, {
					specialEasing: {},
					easing: jQuery.easing._default
				}, options ),
				originalProperties: properties,
				originalOptions: options,
				startTime: fxNow || createFxNow(),
				duration: options.duration,
				tweens: [],
				createTween: function( prop, end ) {
					var tween = jQuery.Tween( elem, animation.opts, prop, end,
							animation.opts.specialEasing[ prop ] || animation.opts.easing );
					animation.tweens.push( tween );
					return tween;
				},
				stop: function( gotoEnd ) {
					var index = 0,
	
						// If we are going to the end, we want to run all the tweens
						// otherwise we skip this part
						length = gotoEnd ? animation.tweens.length : 0;
					if ( stopped ) {
						return this;
					}
					stopped = true;
					for ( ; index < length; index++ ) {
						animation.tweens[ index ].run( 1 );
					}
	
					// Resolve when we played the last frame; otherwise, reject
					if ( gotoEnd ) {
						deferred.notifyWith( elem, [ animation, 1, 0 ] );
						deferred.resolveWith( elem, [ animation, gotoEnd ] );
					} else {
						deferred.rejectWith( elem, [ animation, gotoEnd ] );
					}
					return this;
				}
			} ),
			props = animation.props;
	
		propFilter( props, animation.opts.specialEasing );
	
		for ( ; index < length; index++ ) {
			result = Animation.prefilters[ index ].call( animation, elem, props, animation.opts );
			if ( result ) {
				if ( jQuery.isFunction( result.stop ) ) {
					jQuery._queueHooks( animation.elem, animation.opts.queue ).stop =
						jQuery.proxy( result.stop, result );
				}
				return result;
			}
		}
	
		jQuery.map( props, createTween, animation );
	
		if ( jQuery.isFunction( animation.opts.start ) ) {
			animation.opts.start.call( elem, animation );
		}
	
		// Attach callbacks from options
		animation
			.progress( animation.opts.progress )
			.done( animation.opts.done, animation.opts.complete )
			.fail( animation.opts.fail )
			.always( animation.opts.always );
	
		jQuery.fx.timer(
			jQuery.extend( tick, {
				elem: elem,
				anim: animation,
				queue: animation.opts.queue
			} )
		);
	
		return animation;
	}
	
	jQuery.Animation = jQuery.extend( Animation, {
	
		tweeners: {
			"*": [ function( prop, value ) {
				var tween = this.createTween( prop, value );
				adjustCSS( tween.elem, prop, rcssNum.exec( value ), tween );
				return tween;
			} ]
		},
	
		tweener: function( props, callback ) {
			if ( jQuery.isFunction( props ) ) {
				callback = props;
				props = [ "*" ];
			} else {
				props = props.match( rnothtmlwhite );
			}
	
			var prop,
				index = 0,
				length = props.length;
	
			for ( ; index < length; index++ ) {
				prop = props[ index ];
				Animation.tweeners[ prop ] = Animation.tweeners[ prop ] || [];
				Animation.tweeners[ prop ].unshift( callback );
			}
		},
	
		prefilters: [ defaultPrefilter ],
	
		prefilter: function( callback, prepend ) {
			if ( prepend ) {
				Animation.prefilters.unshift( callback );
			} else {
				Animation.prefilters.push( callback );
			}
		}
	} );
	
	jQuery.speed = function( speed, easing, fn ) {
		var opt = speed && typeof speed === "object" ? jQuery.extend( {}, speed ) : {
			complete: fn || !fn && easing ||
				jQuery.isFunction( speed ) && speed,
			duration: speed,
			easing: fn && easing || easing && !jQuery.isFunction( easing ) && easing
		};
	
		// Go to the end state if fx are off
		if ( jQuery.fx.off ) {
			opt.duration = 0;
	
		} else {
			if ( typeof opt.duration !== "number" ) {
				if ( opt.duration in jQuery.fx.speeds ) {
					opt.duration = jQuery.fx.speeds[ opt.duration ];
	
				} else {
					opt.duration = jQuery.fx.speeds._default;
				}
			}
		}
	
		// Normalize opt.queue - true/undefined/null -> "fx"
		if ( opt.queue == null || opt.queue === true ) {
			opt.queue = "fx";
		}
	
		// Queueing
		opt.old = opt.complete;
	
		opt.complete = function() {
			if ( jQuery.isFunction( opt.old ) ) {
				opt.old.call( this );
			}
	
			if ( opt.queue ) {
				jQuery.dequeue( this, opt.queue );
			}
		};
	
		return opt;
	};
	
	jQuery.fn.extend( {
		fadeTo: function( speed, to, easing, callback ) {
	
			// Show any hidden elements after setting opacity to 0
			return this.filter( isHiddenWithinTree ).css( "opacity", 0 ).show()
	
				// Animate to the value specified
				.end().animate( { opacity: to }, speed, easing, callback );
		},
		animate: function( prop, speed, easing, callback ) {
			var empty = jQuery.isEmptyObject( prop ),
				optall = jQuery.speed( speed, easing, callback ),
				doAnimation = function() {
	
					// Operate on a copy of prop so per-property easing won't be lost
					var anim = Animation( this, jQuery.extend( {}, prop ), optall );
	
					// Empty animations, or finishing resolves immediately
					if ( empty || dataPriv.get( this, "finish" ) ) {
						anim.stop( true );
					}
				};
				doAnimation.finish = doAnimation;
	
			return empty || optall.queue === false ?
				this.each( doAnimation ) :
				this.queue( optall.queue, doAnimation );
		},
		stop: function( type, clearQueue, gotoEnd ) {
			var stopQueue = function( hooks ) {
				var stop = hooks.stop;
				delete hooks.stop;
				stop( gotoEnd );
			};
	
			if ( typeof type !== "string" ) {
				gotoEnd = clearQueue;
				clearQueue = type;
				type = undefined;
			}
			if ( clearQueue && type !== false ) {
				this.queue( type || "fx", [] );
			}
	
			return this.each( function() {
				var dequeue = true,
					index = type != null && type + "queueHooks",
					timers = jQuery.timers,
					data = dataPriv.get( this );
	
				if ( index ) {
					if ( data[ index ] && data[ index ].stop ) {
						stopQueue( data[ index ] );
					}
				} else {
					for ( index in data ) {
						if ( data[ index ] && data[ index ].stop && rrun.test( index ) ) {
							stopQueue( data[ index ] );
						}
					}
				}
	
				for ( index = timers.length; index--; ) {
					if ( timers[ index ].elem === this &&
						( type == null || timers[ index ].queue === type ) ) {
	
						timers[ index ].anim.stop( gotoEnd );
						dequeue = false;
						timers.splice( index, 1 );
					}
				}
	
				// Start the next in the queue if the last step wasn't forced.
				// Timers currently will call their complete callbacks, which
				// will dequeue but only if they were gotoEnd.
				if ( dequeue || !gotoEnd ) {
					jQuery.dequeue( this, type );
				}
			} );
		},
		finish: function( type ) {
			if ( type !== false ) {
				type = type || "fx";
			}
			return this.each( function() {
				var index,
					data = dataPriv.get( this ),
					queue = data[ type + "queue" ],
					hooks = data[ type + "queueHooks" ],
					timers = jQuery.timers,
					length = queue ? queue.length : 0;
	
				// Enable finishing flag on private data
				data.finish = true;
	
				// Empty the queue first
				jQuery.queue( this, type, [] );
	
				if ( hooks && hooks.stop ) {
					hooks.stop.call( this, true );
				}
	
				// Look for any active animations, and finish them
				for ( index = timers.length; index--; ) {
					if ( timers[ index ].elem === this && timers[ index ].queue === type ) {
						timers[ index ].anim.stop( true );
						timers.splice( index, 1 );
					}
				}
	
				// Look for any animations in the old queue and finish them
				for ( index = 0; index < length; index++ ) {
					if ( queue[ index ] && queue[ index ].finish ) {
						queue[ index ].finish.call( this );
					}
				}
	
				// Turn off finishing flag
				delete data.finish;
			} );
		}
	} );
	
	jQuery.each( [ "toggle", "show", "hide" ], function( i, name ) {
		var cssFn = jQuery.fn[ name ];
		jQuery.fn[ name ] = function( speed, easing, callback ) {
			return speed == null || typeof speed === "boolean" ?
				cssFn.apply( this, arguments ) :
				this.animate( genFx( name, true ), speed, easing, callback );
		};
	} );
	
	// Generate shortcuts for custom animations
	jQuery.each( {
		slideDown: genFx( "show" ),
		slideUp: genFx( "hide" ),
		slideToggle: genFx( "toggle" ),
		fadeIn: { opacity: "show" },
		fadeOut: { opacity: "hide" },
		fadeToggle: { opacity: "toggle" }
	}, function( name, props ) {
		jQuery.fn[ name ] = function( speed, easing, callback ) {
			return this.animate( props, speed, easing, callback );
		};
	} );
	
	jQuery.timers = [];
	jQuery.fx.tick = function() {
		var timer,
			i = 0,
			timers = jQuery.timers;
	
		fxNow = jQuery.now();
	
		for ( ; i < timers.length; i++ ) {
			timer = timers[ i ];
	
			// Run the timer and safely remove it when done (allowing for external removal)
			if ( !timer() && timers[ i ] === timer ) {
				timers.splice( i--, 1 );
			}
		}
	
		if ( !timers.length ) {
			jQuery.fx.stop();
		}
		fxNow = undefined;
	};
	
	jQuery.fx.timer = function( timer ) {
		jQuery.timers.push( timer );
		jQuery.fx.start();
	};
	
	jQuery.fx.interval = 13;
	jQuery.fx.start = function() {
		if ( inProgress ) {
			return;
		}
	
		inProgress = true;
		schedule();
	};
	
	jQuery.fx.stop = function() {
		inProgress = null;
	};
	
	jQuery.fx.speeds = {
		slow: 600,
		fast: 200,
	
		// Default speed
		_default: 400
	};
	
	
	// Based off of the plugin by Clint Helfers, with permission.
	// https://web.archive.org/web/20100324014747/http://blindsignals.com/index.php/2009/07/jquery-delay/
	jQuery.fn.delay = function( time, type ) {
		time = jQuery.fx ? jQuery.fx.speeds[ time ] || time : time;
		type = type || "fx";
	
		return this.queue( type, function( next, hooks ) {
			var timeout = window.setTimeout( next, time );
			hooks.stop = function() {
				window.clearTimeout( timeout );
			};
		} );
	};
	
	
	( function() {
		var input = document.createElement( "input" ),
			select = document.createElement( "select" ),
			opt = select.appendChild( document.createElement( "option" ) );
	
		input.type = "checkbox";
	
		// Support: Android <=4.3 only
		// Default value for a checkbox should be "on"
		support.checkOn = input.value !== "";
	
		// Support: IE <=11 only
		// Must access selectedIndex to make default options select
		support.optSelected = opt.selected;
	
		// Support: IE <=11 only
		// An input loses its value after becoming a radio
		input = document.createElement( "input" );
		input.value = "t";
		input.type = "radio";
		support.radioValue = input.value === "t";
	} )();
	
	
	var boolHook,
		attrHandle = jQuery.expr.attrHandle;
	
	jQuery.fn.extend( {
		attr: function( name, value ) {
			return access( this, jQuery.attr, name, value, arguments.length > 1 );
		},
	
		removeAttr: function( name ) {
			return this.each( function() {
				jQuery.removeAttr( this, name );
			} );
		}
	} );
	
	jQuery.extend( {
		attr: function( elem, name, value ) {
			var ret, hooks,
				nType = elem.nodeType;
	
			// Don't get/set attributes on text, comment and attribute nodes
			if ( nType === 3 || nType === 8 || nType === 2 ) {
				return;
			}
	
			// Fallback to prop when attributes are not supported
			if ( typeof elem.getAttribute === "undefined" ) {
				return jQuery.prop( elem, name, value );
			}
	
			// Attribute hooks are determined by the lowercase version
			// Grab necessary hook if one is defined
			if ( nType !== 1 || !jQuery.isXMLDoc( elem ) ) {
				hooks = jQuery.attrHooks[ name.toLowerCase() ] ||
					( jQuery.expr.match.bool.test( name ) ? boolHook : undefined );
			}
	
			if ( value !== undefined ) {
				if ( value === null ) {
					jQuery.removeAttr( elem, name );
					return;
				}
	
				if ( hooks && "set" in hooks &&
					( ret = hooks.set( elem, value, name ) ) !== undefined ) {
					return ret;
				}
	
				elem.setAttribute( name, value + "" );
				return value;
			}
	
			if ( hooks && "get" in hooks && ( ret = hooks.get( elem, name ) ) !== null ) {
				return ret;
			}
	
			ret = jQuery.find.attr( elem, name );
	
			// Non-existent attributes return null, we normalize to undefined
			return ret == null ? undefined : ret;
		},
	
		attrHooks: {
			type: {
				set: function( elem, value ) {
					if ( !support.radioValue && value === "radio" &&
						nodeName( elem, "input" ) ) {
						var val = elem.value;
						elem.setAttribute( "type", value );
						if ( val ) {
							elem.value = val;
						}
						return value;
					}
				}
			}
		},
	
		removeAttr: function( elem, value ) {
			var name,
				i = 0,
	
				// Attribute names can contain non-HTML whitespace characters
				// https://html.spec.whatwg.org/multipage/syntax.html#attributes-2
				attrNames = value && value.match( rnothtmlwhite );
	
			if ( attrNames && elem.nodeType === 1 ) {
				while ( ( name = attrNames[ i++ ] ) ) {
					elem.removeAttribute( name );
				}
			}
		}
	} );
	
	// Hooks for boolean attributes
	boolHook = {
		set: function( elem, value, name ) {
			if ( value === false ) {
	
				// Remove boolean attributes when set to false
				jQuery.removeAttr( elem, name );
			} else {
				elem.setAttribute( name, name );
			}
			return name;
		}
	};
	
	jQuery.each( jQuery.expr.match.bool.source.match( /\w+/g ), function( i, name ) {
		var getter = attrHandle[ name ] || jQuery.find.attr;
	
		attrHandle[ name ] = function( elem, name, isXML ) {
			var ret, handle,
				lowercaseName = name.toLowerCase();
	
			if ( !isXML ) {
	
				// Avoid an infinite loop by temporarily removing this function from the getter
				handle = attrHandle[ lowercaseName ];
				attrHandle[ lowercaseName ] = ret;
				ret = getter( elem, name, isXML ) != null ?
					lowercaseName :
					null;
				attrHandle[ lowercaseName ] = handle;
			}
			return ret;
		};
	} );
	
	
	
	
	var rfocusable = /^(?:input|select|textarea|button)$/i,
		rclickable = /^(?:a|area)$/i;
	
	jQuery.fn.extend( {
		prop: function( name, value ) {
			return access( this, jQuery.prop, name, value, arguments.length > 1 );
		},
	
		removeProp: function( name ) {
			return this.each( function() {
				delete this[ jQuery.propFix[ name ] || name ];
			} );
		}
	} );
	
	jQuery.extend( {
		prop: function( elem, name, value ) {
			var ret, hooks,
				nType = elem.nodeType;
	
			// Don't get/set properties on text, comment and attribute nodes
			if ( nType === 3 || nType === 8 || nType === 2 ) {
				return;
			}
	
			if ( nType !== 1 || !jQuery.isXMLDoc( elem ) ) {
	
				// Fix name and attach hooks
				name = jQuery.propFix[ name ] || name;
				hooks = jQuery.propHooks[ name ];
			}
	
			if ( value !== undefined ) {
				if ( hooks && "set" in hooks &&
					( ret = hooks.set( elem, value, name ) ) !== undefined ) {
					return ret;
				}
	
				return ( elem[ name ] = value );
			}
	
			if ( hooks && "get" in hooks && ( ret = hooks.get( elem, name ) ) !== null ) {
				return ret;
			}
	
			return elem[ name ];
		},
	
		propHooks: {
			tabIndex: {
				get: function( elem ) {
	
					// Support: IE <=9 - 11 only
					// elem.tabIndex doesn't always return the
					// correct value when it hasn't been explicitly set
					// https://web.archive.org/web/20141116233347/http://fluidproject.org/blog/2008/01/09/getting-setting-and-removing-tabindex-values-with-javascript/
					// Use proper attribute retrieval(#12072)
					var tabindex = jQuery.find.attr( elem, "tabindex" );
	
					if ( tabindex ) {
						return parseInt( tabindex, 10 );
					}
	
					if (
						rfocusable.test( elem.nodeName ) ||
						rclickable.test( elem.nodeName ) &&
						elem.href
					) {
						return 0;
					}
	
					return -1;
				}
			}
		},
	
		propFix: {
			"for": "htmlFor",
			"class": "className"
		}
	} );
	
	// Support: IE <=11 only
	// Accessing the selectedIndex property
	// forces the browser to respect setting selected
	// on the option
	// The getter ensures a default option is selected
	// when in an optgroup
	// eslint rule "no-unused-expressions" is disabled for this code
	// since it considers such accessions noop
	if ( !support.optSelected ) {
		jQuery.propHooks.selected = {
			get: function( elem ) {
	
				/* eslint no-unused-expressions: "off" */
	
				var parent = elem.parentNode;
				if ( parent && parent.parentNode ) {
					parent.parentNode.selectedIndex;
				}
				return null;
			},
			set: function( elem ) {
	
				/* eslint no-unused-expressions: "off" */
	
				var parent = elem.parentNode;
				if ( parent ) {
					parent.selectedIndex;
	
					if ( parent.parentNode ) {
						parent.parentNode.selectedIndex;
					}
				}
			}
		};
	}
	
	jQuery.each( [
		"tabIndex",
		"readOnly",
		"maxLength",
		"cellSpacing",
		"cellPadding",
		"rowSpan",
		"colSpan",
		"useMap",
		"frameBorder",
		"contentEditable"
	], function() {
		jQuery.propFix[ this.toLowerCase() ] = this;
	} );
	
	
	
	
		// Strip and collapse whitespace according to HTML spec
		// https://html.spec.whatwg.org/multipage/infrastructure.html#strip-and-collapse-whitespace
		function stripAndCollapse( value ) {
			var tokens = value.match( rnothtmlwhite ) || [];
			return tokens.join( " " );
		}
	
	
	function getClass( elem ) {
		return elem.getAttribute && elem.getAttribute( "class" ) || "";
	}
	
	jQuery.fn.extend( {
		addClass: function( value ) {
			var classes, elem, cur, curValue, clazz, j, finalValue,
				i = 0;
	
			if ( jQuery.isFunction( value ) ) {
				return this.each( function( j ) {
					jQuery( this ).addClass( value.call( this, j, getClass( this ) ) );
				} );
			}
	
			if ( typeof value === "string" && value ) {
				classes = value.match( rnothtmlwhite ) || [];
	
				while ( ( elem = this[ i++ ] ) ) {
					curValue = getClass( elem );
					cur = elem.nodeType === 1 && ( " " + stripAndCollapse( curValue ) + " " );
	
					if ( cur ) {
						j = 0;
						while ( ( clazz = classes[ j++ ] ) ) {
							if ( cur.indexOf( " " + clazz + " " ) < 0 ) {
								cur += clazz + " ";
							}
						}
	
						// Only assign if different to avoid unneeded rendering.
						finalValue = stripAndCollapse( cur );
						if ( curValue !== finalValue ) {
							elem.setAttribute( "class", finalValue );
						}
					}
				}
			}
	
			return this;
		},
	
		removeClass: function( value ) {
			var classes, elem, cur, curValue, clazz, j, finalValue,
				i = 0;
	
			if ( jQuery.isFunction( value ) ) {
				return this.each( function( j ) {
					jQuery( this ).removeClass( value.call( this, j, getClass( this ) ) );
				} );
			}
	
			if ( !arguments.length ) {
				return this.attr( "class", "" );
			}
	
			if ( typeof value === "string" && value ) {
				classes = value.match( rnothtmlwhite ) || [];
	
				while ( ( elem = this[ i++ ] ) ) {
					curValue = getClass( elem );
	
					// This expression is here for better compressibility (see addClass)
					cur = elem.nodeType === 1 && ( " " + stripAndCollapse( curValue ) + " " );
	
					if ( cur ) {
						j = 0;
						while ( ( clazz = classes[ j++ ] ) ) {
	
							// Remove *all* instances
							while ( cur.indexOf( " " + clazz + " " ) > -1 ) {
								cur = cur.replace( " " + clazz + " ", " " );
							}
						}
	
						// Only assign if different to avoid unneeded rendering.
						finalValue = stripAndCollapse( cur );
						if ( curValue !== finalValue ) {
							elem.setAttribute( "class", finalValue );
						}
					}
				}
			}
	
			return this;
		},
	
		toggleClass: function( value, stateVal ) {
			var type = typeof value;
	
			if ( typeof stateVal === "boolean" && type === "string" ) {
				return stateVal ? this.addClass( value ) : this.removeClass( value );
			}
	
			if ( jQuery.isFunction( value ) ) {
				return this.each( function( i ) {
					jQuery( this ).toggleClass(
						value.call( this, i, getClass( this ), stateVal ),
						stateVal
					);
				} );
			}
	
			return this.each( function() {
				var className, i, self, classNames;
	
				if ( type === "string" ) {
	
					// Toggle individual class names
					i = 0;
					self = jQuery( this );
					classNames = value.match( rnothtmlwhite ) || [];
	
					while ( ( className = classNames[ i++ ] ) ) {
	
						// Check each className given, space separated list
						if ( self.hasClass( className ) ) {
							self.removeClass( className );
						} else {
							self.addClass( className );
						}
					}
	
				// Toggle whole class name
				} else if ( value === undefined || type === "boolean" ) {
					className = getClass( this );
					if ( className ) {
	
						// Store className if set
						dataPriv.set( this, "__className__", className );
					}
	
					// If the element has a class name or if we're passed `false`,
					// then remove the whole classname (if there was one, the above saved it).
					// Otherwise bring back whatever was previously saved (if anything),
					// falling back to the empty string if nothing was stored.
					if ( this.setAttribute ) {
						this.setAttribute( "class",
							className || value === false ?
							"" :
							dataPriv.get( this, "__className__" ) || ""
						);
					}
				}
			} );
		},
	
		hasClass: function( selector ) {
			var className, elem,
				i = 0;
	
			className = " " + selector + " ";
			while ( ( elem = this[ i++ ] ) ) {
				if ( elem.nodeType === 1 &&
					( " " + stripAndCollapse( getClass( elem ) ) + " " ).indexOf( className ) > -1 ) {
						return true;
				}
			}
	
			return false;
		}
	} );
	
	
	
	
	var rreturn = /\r/g;
	
	jQuery.fn.extend( {
		val: function( value ) {
			var hooks, ret, isFunction,
				elem = this[ 0 ];
	
			if ( !arguments.length ) {
				if ( elem ) {
					hooks = jQuery.valHooks[ elem.type ] ||
						jQuery.valHooks[ elem.nodeName.toLowerCase() ];
	
					if ( hooks &&
						"get" in hooks &&
						( ret = hooks.get( elem, "value" ) ) !== undefined
					) {
						return ret;
					}
	
					ret = elem.value;
	
					// Handle most common string cases
					if ( typeof ret === "string" ) {
						return ret.replace( rreturn, "" );
					}
	
					// Handle cases where value is null/undef or number
					return ret == null ? "" : ret;
				}
	
				return;
			}
	
			isFunction = jQuery.isFunction( value );
	
			return this.each( function( i ) {
				var val;
	
				if ( this.nodeType !== 1 ) {
					return;
				}
	
				if ( isFunction ) {
					val = value.call( this, i, jQuery( this ).val() );
				} else {
					val = value;
				}
	
				// Treat null/undefined as ""; convert numbers to string
				if ( val == null ) {
					val = "";
	
				} else if ( typeof val === "number" ) {
					val += "";
	
				} else if ( Array.isArray( val ) ) {
					val = jQuery.map( val, function( value ) {
						return value == null ? "" : value + "";
					} );
				}
	
				hooks = jQuery.valHooks[ this.type ] || jQuery.valHooks[ this.nodeName.toLowerCase() ];
	
				// If set returns undefined, fall back to normal setting
				if ( !hooks || !( "set" in hooks ) || hooks.set( this, val, "value" ) === undefined ) {
					this.value = val;
				}
			} );
		}
	} );
	
	jQuery.extend( {
		valHooks: {
			option: {
				get: function( elem ) {
	
					var val = jQuery.find.attr( elem, "value" );
					return val != null ?
						val :
	
						// Support: IE <=10 - 11 only
						// option.text throws exceptions (#14686, #14858)
						// Strip and collapse whitespace
						// https://html.spec.whatwg.org/#strip-and-collapse-whitespace
						stripAndCollapse( jQuery.text( elem ) );
				}
			},
			select: {
				get: function( elem ) {
					var value, option, i,
						options = elem.options,
						index = elem.selectedIndex,
						one = elem.type === "select-one",
						values = one ? null : [],
						max = one ? index + 1 : options.length;
	
					if ( index < 0 ) {
						i = max;
	
					} else {
						i = one ? index : 0;
					}
	
					// Loop through all the selected options
					for ( ; i < max; i++ ) {
						option = options[ i ];
	
						// Support: IE <=9 only
						// IE8-9 doesn't update selected after form reset (#2551)
						if ( ( option.selected || i === index ) &&
	
								// Don't return options that are disabled or in a disabled optgroup
								!option.disabled &&
								( !option.parentNode.disabled ||
									!nodeName( option.parentNode, "optgroup" ) ) ) {
	
							// Get the specific value for the option
							value = jQuery( option ).val();
	
							// We don't need an array for one selects
							if ( one ) {
								return value;
							}
	
							// Multi-Selects return an array
							values.push( value );
						}
					}
	
					return values;
				},
	
				set: function( elem, value ) {
					var optionSet, option,
						options = elem.options,
						values = jQuery.makeArray( value ),
						i = options.length;
	
					while ( i-- ) {
						option = options[ i ];
	
						/* eslint-disable no-cond-assign */
	
						if ( option.selected =
							jQuery.inArray( jQuery.valHooks.option.get( option ), values ) > -1
						) {
							optionSet = true;
						}
	
						/* eslint-enable no-cond-assign */
					}
	
					// Force browsers to behave consistently when non-matching value is set
					if ( !optionSet ) {
						elem.selectedIndex = -1;
					}
					return values;
				}
			}
		}
	} );
	
	// Radios and checkboxes getter/setter
	jQuery.each( [ "radio", "checkbox" ], function() {
		jQuery.valHooks[ this ] = {
			set: function( elem, value ) {
				if ( Array.isArray( value ) ) {
					return ( elem.checked = jQuery.inArray( jQuery( elem ).val(), value ) > -1 );
				}
			}
		};
		if ( !support.checkOn ) {
			jQuery.valHooks[ this ].get = function( elem ) {
				return elem.getAttribute( "value" ) === null ? "on" : elem.value;
			};
		}
	} );
	
	
	
	
	// Return jQuery for attributes-only inclusion
	
	
	var rfocusMorph = /^(?:focusinfocus|focusoutblur)$/;
	
	jQuery.extend( jQuery.event, {
	
		trigger: function( event, data, elem, onlyHandlers ) {
	
			var i, cur, tmp, bubbleType, ontype, handle, special,
				eventPath = [ elem || document ],
				type = hasOwn.call( event, "type" ) ? event.type : event,
				namespaces = hasOwn.call( event, "namespace" ) ? event.namespace.split( "." ) : [];
	
			cur = tmp = elem = elem || document;
	
			// Don't do events on text and comment nodes
			if ( elem.nodeType === 3 || elem.nodeType === 8 ) {
				return;
			}
	
			// focus/blur morphs to focusin/out; ensure we're not firing them right now
			if ( rfocusMorph.test( type + jQuery.event.triggered ) ) {
				return;
			}
	
			if ( type.indexOf( "." ) > -1 ) {
	
				// Namespaced trigger; create a regexp to match event type in handle()
				namespaces = type.split( "." );
				type = namespaces.shift();
				namespaces.sort();
			}
			ontype = type.indexOf( ":" ) < 0 && "on" + type;
	
			// Caller can pass in a jQuery.Event object, Object, or just an event type string
			event = event[ jQuery.expando ] ?
				event :
				new jQuery.Event( type, typeof event === "object" && event );
	
			// Trigger bitmask: & 1 for native handlers; & 2 for jQuery (always true)
			event.isTrigger = onlyHandlers ? 2 : 3;
			event.namespace = namespaces.join( "." );
			event.rnamespace = event.namespace ?
				new RegExp( "(^|\\.)" + namespaces.join( "\\.(?:.*\\.|)" ) + "(\\.|$)" ) :
				null;
	
			// Clean up the event in case it is being reused
			event.result = undefined;
			if ( !event.target ) {
				event.target = elem;
			}
	
			// Clone any incoming data and prepend the event, creating the handler arg list
			data = data == null ?
				[ event ] :
				jQuery.makeArray( data, [ event ] );
	
			// Allow special events to draw outside the lines
			special = jQuery.event.special[ type ] || {};
			if ( !onlyHandlers && special.trigger && special.trigger.apply( elem, data ) === false ) {
				return;
			}
	
			// Determine event propagation path in advance, per W3C events spec (#9951)
			// Bubble up to document, then to window; watch for a global ownerDocument var (#9724)
			if ( !onlyHandlers && !special.noBubble && !jQuery.isWindow( elem ) ) {
	
				bubbleType = special.delegateType || type;
				if ( !rfocusMorph.test( bubbleType + type ) ) {
					cur = cur.parentNode;
				}
				for ( ; cur; cur = cur.parentNode ) {
					eventPath.push( cur );
					tmp = cur;
				}
	
				// Only add window if we got to document (e.g., not plain obj or detached DOM)
				if ( tmp === ( elem.ownerDocument || document ) ) {
					eventPath.push( tmp.defaultView || tmp.parentWindow || window );
				}
			}
	
			// Fire handlers on the event path
			i = 0;
			while ( ( cur = eventPath[ i++ ] ) && !event.isPropagationStopped() ) {
	
				event.type = i > 1 ?
					bubbleType :
					special.bindType || type;
	
				// jQuery handler
				handle = ( dataPriv.get( cur, "events" ) || {} )[ event.type ] &&
					dataPriv.get( cur, "handle" );
				if ( handle ) {
					handle.apply( cur, data );
				}
	
				// Native handler
				handle = ontype && cur[ ontype ];
				if ( handle && handle.apply && acceptData( cur ) ) {
					event.result = handle.apply( cur, data );
					if ( event.result === false ) {
						event.preventDefault();
					}
				}
			}
			event.type = type;
	
			// If nobody prevented the default action, do it now
			if ( !onlyHandlers && !event.isDefaultPrevented() ) {
	
				if ( ( !special._default ||
					special._default.apply( eventPath.pop(), data ) === false ) &&
					acceptData( elem ) ) {
	
					// Call a native DOM method on the target with the same name as the event.
					// Don't do default actions on window, that's where global variables be (#6170)
					if ( ontype && jQuery.isFunction( elem[ type ] ) && !jQuery.isWindow( elem ) ) {
	
						// Don't re-trigger an onFOO event when we call its FOO() method
						tmp = elem[ ontype ];
	
						if ( tmp ) {
							elem[ ontype ] = null;
						}
	
						// Prevent re-triggering of the same event, since we already bubbled it above
						jQuery.event.triggered = type;
						elem[ type ]();
						jQuery.event.triggered = undefined;
	
						if ( tmp ) {
							elem[ ontype ] = tmp;
						}
					}
				}
			}
	
			return event.result;
		},
	
		// Piggyback on a donor event to simulate a different one
		// Used only for `focus(in | out)` events
		simulate: function( type, elem, event ) {
			var e = jQuery.extend(
				new jQuery.Event(),
				event,
				{
					type: type,
					isSimulated: true
				}
			);
	
			jQuery.event.trigger( e, null, elem );
		}
	
	} );
	
	jQuery.fn.extend( {
	
		trigger: function( type, data ) {
			return this.each( function() {
				jQuery.event.trigger( type, data, this );
			} );
		},
		triggerHandler: function( type, data ) {
			var elem = this[ 0 ];
			if ( elem ) {
				return jQuery.event.trigger( type, data, elem, true );
			}
		}
	} );
	
	
	jQuery.each( ( "blur focus focusin focusout resize scroll click dblclick " +
		"mousedown mouseup mousemove mouseover mouseout mouseenter mouseleave " +
		"change select submit keydown keypress keyup contextmenu" ).split( " " ),
		function( i, name ) {
	
		// Handle event binding
		jQuery.fn[ name ] = function( data, fn ) {
			return arguments.length > 0 ?
				this.on( name, null, data, fn ) :
				this.trigger( name );
		};
	} );
	
	jQuery.fn.extend( {
		hover: function( fnOver, fnOut ) {
			return this.mouseenter( fnOver ).mouseleave( fnOut || fnOver );
		}
	} );
	
	
	
	
	support.focusin = "onfocusin" in window;
	
	
	// Support: Firefox <=44
	// Firefox doesn't have focus(in | out) events
	// Related ticket - https://bugzilla.mozilla.org/show_bug.cgi?id=687787
	//
	// Support: Chrome <=48 - 49, Safari <=9.0 - 9.1
	// focus(in | out) events fire after focus & blur events,
	// which is spec violation - http://www.w3.org/TR/DOM-Level-3-Events/#events-focusevent-event-order
	// Related ticket - https://bugs.chromium.org/p/chromium/issues/detail?id=449857
	if ( !support.focusin ) {
		jQuery.each( { focus: "focusin", blur: "focusout" }, function( orig, fix ) {
	
			// Attach a single capturing handler on the document while someone wants focusin/focusout
			var handler = function( event ) {
				jQuery.event.simulate( fix, event.target, jQuery.event.fix( event ) );
			};
	
			jQuery.event.special[ fix ] = {
				setup: function() {
					var doc = this.ownerDocument || this,
						attaches = dataPriv.access( doc, fix );
	
					if ( !attaches ) {
						doc.addEventListener( orig, handler, true );
					}
					dataPriv.access( doc, fix, ( attaches || 0 ) + 1 );
				},
				teardown: function() {
					var doc = this.ownerDocument || this,
						attaches = dataPriv.access( doc, fix ) - 1;
	
					if ( !attaches ) {
						doc.removeEventListener( orig, handler, true );
						dataPriv.remove( doc, fix );
	
					} else {
						dataPriv.access( doc, fix, attaches );
					}
				}
			};
		} );
	}
	var location = window.location;
	
	var nonce = jQuery.now();
	
	var rquery = ( /\?/ );
	
	
	
	// Cross-browser xml parsing
	jQuery.parseXML = function( data ) {
		var xml;
		if ( !data || typeof data !== "string" ) {
			return null;
		}
	
		// Support: IE 9 - 11 only
		// IE throws on parseFromString with invalid input.
		try {
			xml = ( new window.DOMParser() ).parseFromString( data, "text/xml" );
		} catch ( e ) {
			xml = undefined;
		}
	
		if ( !xml || xml.getElementsByTagName( "parsererror" ).length ) {
			jQuery.error( "Invalid XML: " + data );
		}
		return xml;
	};
	
	
	var
		rbracket = /\[\]$/,
		rCRLF = /\r?\n/g,
		rsubmitterTypes = /^(?:submit|button|image|reset|file)$/i,
		rsubmittable = /^(?:input|select|textarea|keygen)/i;
	
	function buildParams( prefix, obj, traditional, add ) {
		var name;
	
		if ( Array.isArray( obj ) ) {
	
			// Serialize array item.
			jQuery.each( obj, function( i, v ) {
				if ( traditional || rbracket.test( prefix ) ) {
	
					// Treat each array item as a scalar.
					add( prefix, v );
	
				} else {
	
					// Item is non-scalar (array or object), encode its numeric index.
					buildParams(
						prefix + "[" + ( typeof v === "object" && v != null ? i : "" ) + "]",
						v,
						traditional,
						add
					);
				}
			} );
	
		} else if ( !traditional && jQuery.type( obj ) === "object" ) {
	
			// Serialize object item.
			for ( name in obj ) {
				buildParams( prefix + "[" + name + "]", obj[ name ], traditional, add );
			}
	
		} else {
	
			// Serialize scalar item.
			add( prefix, obj );
		}
	}
	
	// Serialize an array of form elements or a set of
	// key/values into a query string
	jQuery.param = function( a, traditional ) {
		var prefix,
			s = [],
			add = function( key, valueOrFunction ) {
	
				// If value is a function, invoke it and use its return value
				var value = jQuery.isFunction( valueOrFunction ) ?
					valueOrFunction() :
					valueOrFunction;
	
				s[ s.length ] = encodeURIComponent( key ) + "=" +
					encodeURIComponent( value == null ? "" : value );
			};
	
		// If an array was passed in, assume that it is an array of form elements.
		if ( Array.isArray( a ) || ( a.jquery && !jQuery.isPlainObject( a ) ) ) {
	
			// Serialize the form elements
			jQuery.each( a, function() {
				add( this.name, this.value );
			} );
	
		} else {
	
			// If traditional, encode the "old" way (the way 1.3.2 or older
			// did it), otherwise encode params recursively.
			for ( prefix in a ) {
				buildParams( prefix, a[ prefix ], traditional, add );
			}
		}
	
		// Return the resulting serialization
		return s.join( "&" );
	};
	
	jQuery.fn.extend( {
		serialize: function() {
			return jQuery.param( this.serializeArray() );
		},
		serializeArray: function() {
			return this.map( function() {
	
				// Can add propHook for "elements" to filter or add form elements
				var elements = jQuery.prop( this, "elements" );
				return elements ? jQuery.makeArray( elements ) : this;
			} )
			.filter( function() {
				var type = this.type;
	
				// Use .is( ":disabled" ) so that fieldset[disabled] works
				return this.name && !jQuery( this ).is( ":disabled" ) &&
					rsubmittable.test( this.nodeName ) && !rsubmitterTypes.test( type ) &&
					( this.checked || !rcheckableType.test( type ) );
			} )
			.map( function( i, elem ) {
				var val = jQuery( this ).val();
	
				if ( val == null ) {
					return null;
				}
	
				if ( Array.isArray( val ) ) {
					return jQuery.map( val, function( val ) {
						return { name: elem.name, value: val.replace( rCRLF, "\r\n" ) };
					} );
				}
	
				return { name: elem.name, value: val.replace( rCRLF, "\r\n" ) };
			} ).get();
		}
	} );
	
	
	var
		r20 = /%20/g,
		rhash = /#.*$/,
		rantiCache = /([?&])_=[^&]*/,
		rheaders = /^(.*?):[ \t]*([^\r\n]*)$/mg,
	
		// #7653, #8125, #8152: local protocol detection
		rlocalProtocol = /^(?:about|app|app-storage|.+-extension|file|res|widget):$/,
		rnoContent = /^(?:GET|HEAD)$/,
		rprotocol = /^\/\//,
	
		/* Prefilters
		 * 1) They are useful to introduce custom dataTypes (see ajax/jsonp.js for an example)
		 * 2) These are called:
		 *    - BEFORE asking for a transport
		 *    - AFTER param serialization (s.data is a string if s.processData is true)
		 * 3) key is the dataType
		 * 4) the catchall symbol "*" can be used
		 * 5) execution will start with transport dataType and THEN continue down to "*" if needed
		 */
		prefilters = {},
	
		/* Transports bindings
		 * 1) key is the dataType
		 * 2) the catchall symbol "*" can be used
		 * 3) selection will start with transport dataType and THEN go to "*" if needed
		 */
		transports = {},
	
		// Avoid comment-prolog char sequence (#10098); must appease lint and evade compression
		allTypes = "*/".concat( "*" ),
	
		// Anchor tag for parsing the document origin
		originAnchor = document.createElement( "a" );
		originAnchor.href = location.href;
	
	// Base "constructor" for jQuery.ajaxPrefilter and jQuery.ajaxTransport
	function addToPrefiltersOrTransports( structure ) {
	
		// dataTypeExpression is optional and defaults to "*"
		return function( dataTypeExpression, func ) {
	
			if ( typeof dataTypeExpression !== "string" ) {
				func = dataTypeExpression;
				dataTypeExpression = "*";
			}
	
			var dataType,
				i = 0,
				dataTypes = dataTypeExpression.toLowerCase().match( rnothtmlwhite ) || [];
	
			if ( jQuery.isFunction( func ) ) {
	
				// For each dataType in the dataTypeExpression
				while ( ( dataType = dataTypes[ i++ ] ) ) {
	
					// Prepend if requested
					if ( dataType[ 0 ] === "+" ) {
						dataType = dataType.slice( 1 ) || "*";
						( structure[ dataType ] = structure[ dataType ] || [] ).unshift( func );
	
					// Otherwise append
					} else {
						( structure[ dataType ] = structure[ dataType ] || [] ).push( func );
					}
				}
			}
		};
	}
	
	// Base inspection function for prefilters and transports
	function inspectPrefiltersOrTransports( structure, options, originalOptions, jqXHR ) {
	
		var inspected = {},
			seekingTransport = ( structure === transports );
	
		function inspect( dataType ) {
			var selected;
			inspected[ dataType ] = true;
			jQuery.each( structure[ dataType ] || [], function( _, prefilterOrFactory ) {
				var dataTypeOrTransport = prefilterOrFactory( options, originalOptions, jqXHR );
				if ( typeof dataTypeOrTransport === "string" &&
					!seekingTransport && !inspected[ dataTypeOrTransport ] ) {
	
					options.dataTypes.unshift( dataTypeOrTransport );
					inspect( dataTypeOrTransport );
					return false;
				} else if ( seekingTransport ) {
					return !( selected = dataTypeOrTransport );
				}
			} );
			return selected;
		}
	
		return inspect( options.dataTypes[ 0 ] ) || !inspected[ "*" ] && inspect( "*" );
	}
	
	// A special extend for ajax options
	// that takes "flat" options (not to be deep extended)
	// Fixes #9887
	function ajaxExtend( target, src ) {
		var key, deep,
			flatOptions = jQuery.ajaxSettings.flatOptions || {};
	
		for ( key in src ) {
			if ( src[ key ] !== undefined ) {
				( flatOptions[ key ] ? target : ( deep || ( deep = {} ) ) )[ key ] = src[ key ];
			}
		}
		if ( deep ) {
			jQuery.extend( true, target, deep );
		}
	
		return target;
	}
	
	/* Handles responses to an ajax request:
	 * - finds the right dataType (mediates between content-type and expected dataType)
	 * - returns the corresponding response
	 */
	function ajaxHandleResponses( s, jqXHR, responses ) {
	
		var ct, type, finalDataType, firstDataType,
			contents = s.contents,
			dataTypes = s.dataTypes;
	
		// Remove auto dataType and get content-type in the process
		while ( dataTypes[ 0 ] === "*" ) {
			dataTypes.shift();
			if ( ct === undefined ) {
				ct = s.mimeType || jqXHR.getResponseHeader( "Content-Type" );
			}
		}
	
		// Check if we're dealing with a known content-type
		if ( ct ) {
			for ( type in contents ) {
				if ( contents[ type ] && contents[ type ].test( ct ) ) {
					dataTypes.unshift( type );
					break;
				}
			}
		}
	
		// Check to see if we have a response for the expected dataType
		if ( dataTypes[ 0 ] in responses ) {
			finalDataType = dataTypes[ 0 ];
		} else {
	
			// Try convertible dataTypes
			for ( type in responses ) {
				if ( !dataTypes[ 0 ] || s.converters[ type + " " + dataTypes[ 0 ] ] ) {
					finalDataType = type;
					break;
				}
				if ( !firstDataType ) {
					firstDataType = type;
				}
			}
	
			// Or just use first one
			finalDataType = finalDataType || firstDataType;
		}
	
		// If we found a dataType
		// We add the dataType to the list if needed
		// and return the corresponding response
		if ( finalDataType ) {
			if ( finalDataType !== dataTypes[ 0 ] ) {
				dataTypes.unshift( finalDataType );
			}
			return responses[ finalDataType ];
		}
	}
	
	/* Chain conversions given the request and the original response
	 * Also sets the responseXXX fields on the jqXHR instance
	 */
	function ajaxConvert( s, response, jqXHR, isSuccess ) {
		var conv2, current, conv, tmp, prev,
			converters = {},
	
			// Work with a copy of dataTypes in case we need to modify it for conversion
			dataTypes = s.dataTypes.slice();
	
		// Create converters map with lowercased keys
		if ( dataTypes[ 1 ] ) {
			for ( conv in s.converters ) {
				converters[ conv.toLowerCase() ] = s.converters[ conv ];
			}
		}
	
		current = dataTypes.shift();
	
		// Convert to each sequential dataType
		while ( current ) {
	
			if ( s.responseFields[ current ] ) {
				jqXHR[ s.responseFields[ current ] ] = response;
			}
	
			// Apply the dataFilter if provided
			if ( !prev && isSuccess && s.dataFilter ) {
				response = s.dataFilter( response, s.dataType );
			}
	
			prev = current;
			current = dataTypes.shift();
	
			if ( current ) {
	
				// There's only work to do if current dataType is non-auto
				if ( current === "*" ) {
	
					current = prev;
	
				// Convert response if prev dataType is non-auto and differs from current
				} else if ( prev !== "*" && prev !== current ) {
	
					// Seek a direct converter
					conv = converters[ prev + " " + current ] || converters[ "* " + current ];
	
					// If none found, seek a pair
					if ( !conv ) {
						for ( conv2 in converters ) {
	
							// If conv2 outputs current
							tmp = conv2.split( " " );
							if ( tmp[ 1 ] === current ) {
	
								// If prev can be converted to accepted input
								conv = converters[ prev + " " + tmp[ 0 ] ] ||
									converters[ "* " + tmp[ 0 ] ];
								if ( conv ) {
	
									// Condense equivalence converters
									if ( conv === true ) {
										conv = converters[ conv2 ];
	
									// Otherwise, insert the intermediate dataType
									} else if ( converters[ conv2 ] !== true ) {
										current = tmp[ 0 ];
										dataTypes.unshift( tmp[ 1 ] );
									}
									break;
								}
							}
						}
					}
	
					// Apply converter (if not an equivalence)
					if ( conv !== true ) {
	
						// Unless errors are allowed to bubble, catch and return them
						if ( conv && s.throws ) {
							response = conv( response );
						} else {
							try {
								response = conv( response );
							} catch ( e ) {
								return {
									state: "parsererror",
									error: conv ? e : "No conversion from " + prev + " to " + current
								};
							}
						}
					}
				}
			}
		}
	
		return { state: "success", data: response };
	}
	
	jQuery.extend( {
	
		// Counter for holding the number of active queries
		active: 0,
	
		// Last-Modified header cache for next request
		lastModified: {},
		etag: {},
	
		ajaxSettings: {
			url: location.href,
			type: "GET",
			isLocal: rlocalProtocol.test( location.protocol ),
			global: true,
			processData: true,
			async: true,
			contentType: "application/x-www-form-urlencoded; charset=UTF-8",
	
			/*
			timeout: 0,
			data: null,
			dataType: null,
			username: null,
			password: null,
			cache: null,
			throws: false,
			traditional: false,
			headers: {},
			*/
	
			accepts: {
				"*": allTypes,
				text: "text/plain",
				html: "text/html",
				xml: "application/xml, text/xml",
				json: "application/json, text/javascript"
			},
	
			contents: {
				xml: /\bxml\b/,
				html: /\bhtml/,
				json: /\bjson\b/
			},
	
			responseFields: {
				xml: "responseXML",
				text: "responseText",
				json: "responseJSON"
			},
	
			// Data converters
			// Keys separate source (or catchall "*") and destination types with a single space
			converters: {
	
				// Convert anything to text
				"* text": String,
	
				// Text to html (true = no transformation)
				"text html": true,
	
				// Evaluate text as a json expression
				"text json": JSON.parse,
	
				// Parse text as xml
				"text xml": jQuery.parseXML
			},
	
			// For options that shouldn't be deep extended:
			// you can add your own custom options here if
			// and when you create one that shouldn't be
			// deep extended (see ajaxExtend)
			flatOptions: {
				url: true,
				context: true
			}
		},
	
		// Creates a full fledged settings object into target
		// with both ajaxSettings and settings fields.
		// If target is omitted, writes into ajaxSettings.
		ajaxSetup: function( target, settings ) {
			return settings ?
	
				// Building a settings object
				ajaxExtend( ajaxExtend( target, jQuery.ajaxSettings ), settings ) :
	
				// Extending ajaxSettings
				ajaxExtend( jQuery.ajaxSettings, target );
		},
	
		ajaxPrefilter: addToPrefiltersOrTransports( prefilters ),
		ajaxTransport: addToPrefiltersOrTransports( transports ),
	
		// Main method
		ajax: function( url, options ) {
	
			// If url is an object, simulate pre-1.5 signature
			if ( typeof url === "object" ) {
				options = url;
				url = undefined;
			}
	
			// Force options to be an object
			options = options || {};
	
			var transport,
	
				// URL without anti-cache param
				cacheURL,
	
				// Response headers
				responseHeadersString,
				responseHeaders,
	
				// timeout handle
				timeoutTimer,
	
				// Url cleanup var
				urlAnchor,
	
				// Request state (becomes false upon send and true upon completion)
				completed,
	
				// To know if global events are to be dispatched
				fireGlobals,
	
				// Loop variable
				i,
	
				// uncached part of the url
				uncached,
	
				// Create the final options object
				s = jQuery.ajaxSetup( {}, options ),
	
				// Callbacks context
				callbackContext = s.context || s,
	
				// Context for global events is callbackContext if it is a DOM node or jQuery collection
				globalEventContext = s.context &&
					( callbackContext.nodeType || callbackContext.jquery ) ?
						jQuery( callbackContext ) :
						jQuery.event,
	
				// Deferreds
				deferred = jQuery.Deferred(),
				completeDeferred = jQuery.Callbacks( "once memory" ),
	
				// Status-dependent callbacks
				statusCode = s.statusCode || {},
	
				// Headers (they are sent all at once)
				requestHeaders = {},
				requestHeadersNames = {},
	
				// Default abort message
				strAbort = "canceled",
	
				// Fake xhr
				jqXHR = {
					readyState: 0,
	
					// Builds headers hashtable if needed
					getResponseHeader: function( key ) {
						var match;
						if ( completed ) {
							if ( !responseHeaders ) {
								responseHeaders = {};
								while ( ( match = rheaders.exec( responseHeadersString ) ) ) {
									responseHeaders[ match[ 1 ].toLowerCase() ] = match[ 2 ];
								}
							}
							match = responseHeaders[ key.toLowerCase() ];
						}
						return match == null ? null : match;
					},
	
					// Raw string
					getAllResponseHeaders: function() {
						return completed ? responseHeadersString : null;
					},
	
					// Caches the header
					setRequestHeader: function( name, value ) {
						if ( completed == null ) {
							name = requestHeadersNames[ name.toLowerCase() ] =
								requestHeadersNames[ name.toLowerCase() ] || name;
							requestHeaders[ name ] = value;
						}
						return this;
					},
	
					// Overrides response content-type header
					overrideMimeType: function( type ) {
						if ( completed == null ) {
							s.mimeType = type;
						}
						return this;
					},
	
					// Status-dependent callbacks
					statusCode: function( map ) {
						var code;
						if ( map ) {
							if ( completed ) {
	
								// Execute the appropriate callbacks
								jqXHR.always( map[ jqXHR.status ] );
							} else {
	
								// Lazy-add the new callbacks in a way that preserves old ones
								for ( code in map ) {
									statusCode[ code ] = [ statusCode[ code ], map[ code ] ];
								}
							}
						}
						return this;
					},
	
					// Cancel the request
					abort: function( statusText ) {
						var finalText = statusText || strAbort;
						if ( transport ) {
							transport.abort( finalText );
						}
						done( 0, finalText );
						return this;
					}
				};
	
			// Attach deferreds
			deferred.promise( jqXHR );
	
			// Add protocol if not provided (prefilters might expect it)
			// Handle falsy url in the settings object (#10093: consistency with old signature)
			// We also use the url parameter if available
			s.url = ( ( url || s.url || location.href ) + "" )
				.replace( rprotocol, location.protocol + "//" );
	
			// Alias method option to type as per ticket #12004
			s.type = options.method || options.type || s.method || s.type;
	
			// Extract dataTypes list
			s.dataTypes = ( s.dataType || "*" ).toLowerCase().match( rnothtmlwhite ) || [ "" ];
	
			// A cross-domain request is in order when the origin doesn't match the current origin.
			if ( s.crossDomain == null ) {
				urlAnchor = document.createElement( "a" );
	
				// Support: IE <=8 - 11, Edge 12 - 13
				// IE throws exception on accessing the href property if url is malformed,
				// e.g. http://example.com:80x/
				try {
					urlAnchor.href = s.url;
	
					// Support: IE <=8 - 11 only
					// Anchor's host property isn't correctly set when s.url is relative
					urlAnchor.href = urlAnchor.href;
					s.crossDomain = originAnchor.protocol + "//" + originAnchor.host !==
						urlAnchor.protocol + "//" + urlAnchor.host;
				} catch ( e ) {
	
					// If there is an error parsing the URL, assume it is crossDomain,
					// it can be rejected by the transport if it is invalid
					s.crossDomain = true;
				}
			}
	
			// Convert data if not already a string
			if ( s.data && s.processData && typeof s.data !== "string" ) {
				s.data = jQuery.param( s.data, s.traditional );
			}
	
			// Apply prefilters
			inspectPrefiltersOrTransports( prefilters, s, options, jqXHR );
	
			// If request was aborted inside a prefilter, stop there
			if ( completed ) {
				return jqXHR;
			}
	
			// We can fire global events as of now if asked to
			// Don't fire events if jQuery.event is undefined in an AMD-usage scenario (#15118)
			fireGlobals = jQuery.event && s.global;
	
			// Watch for a new set of requests
			if ( fireGlobals && jQuery.active++ === 0 ) {
				jQuery.event.trigger( "ajaxStart" );
			}
	
			// Uppercase the type
			s.type = s.type.toUpperCase();
	
			// Determine if request has content
			s.hasContent = !rnoContent.test( s.type );
	
			// Save the URL in case we're toying with the If-Modified-Since
			// and/or If-None-Match header later on
			// Remove hash to simplify url manipulation
			cacheURL = s.url.replace( rhash, "" );
	
			// More options handling for requests with no content
			if ( !s.hasContent ) {
	
				// Remember the hash so we can put it back
				uncached = s.url.slice( cacheURL.length );
	
				// If data is available, append data to url
				if ( s.data ) {
					cacheURL += ( rquery.test( cacheURL ) ? "&" : "?" ) + s.data;
	
					// #9682: remove data so that it's not used in an eventual retry
					delete s.data;
				}
	
				// Add or update anti-cache param if needed
				if ( s.cache === false ) {
					cacheURL = cacheURL.replace( rantiCache, "$1" );
					uncached = ( rquery.test( cacheURL ) ? "&" : "?" ) + "_=" + ( nonce++ ) + uncached;
				}
	
				// Put hash and anti-cache on the URL that will be requested (gh-1732)
				s.url = cacheURL + uncached;
	
			// Change '%20' to '+' if this is encoded form body content (gh-2658)
			} else if ( s.data && s.processData &&
				( s.contentType || "" ).indexOf( "application/x-www-form-urlencoded" ) === 0 ) {
				s.data = s.data.replace( r20, "+" );
			}
	
			// Set the If-Modified-Since and/or If-None-Match header, if in ifModified mode.
			if ( s.ifModified ) {
				if ( jQuery.lastModified[ cacheURL ] ) {
					jqXHR.setRequestHeader( "If-Modified-Since", jQuery.lastModified[ cacheURL ] );
				}
				if ( jQuery.etag[ cacheURL ] ) {
					jqXHR.setRequestHeader( "If-None-Match", jQuery.etag[ cacheURL ] );
				}
			}
	
			// Set the correct header, if data is being sent
			if ( s.data && s.hasContent && s.contentType !== false || options.contentType ) {
				jqXHR.setRequestHeader( "Content-Type", s.contentType );
			}
	
			// Set the Accepts header for the server, depending on the dataType
			jqXHR.setRequestHeader(
				"Accept",
				s.dataTypes[ 0 ] && s.accepts[ s.dataTypes[ 0 ] ] ?
					s.accepts[ s.dataTypes[ 0 ] ] +
						( s.dataTypes[ 0 ] !== "*" ? ", " + allTypes + "; q=0.01" : "" ) :
					s.accepts[ "*" ]
			);
	
			// Check for headers option
			for ( i in s.headers ) {
				jqXHR.setRequestHeader( i, s.headers[ i ] );
			}
	
			// Allow custom headers/mimetypes and early abort
			if ( s.beforeSend &&
				( s.beforeSend.call( callbackContext, jqXHR, s ) === false || completed ) ) {
	
				// Abort if not done already and return
				return jqXHR.abort();
			}
	
			// Aborting is no longer a cancellation
			strAbort = "abort";
	
			// Install callbacks on deferreds
			completeDeferred.add( s.complete );
			jqXHR.done( s.success );
			jqXHR.fail( s.error );
	
			// Get transport
			transport = inspectPrefiltersOrTransports( transports, s, options, jqXHR );
	
			// If no transport, we auto-abort
			if ( !transport ) {
				done( -1, "No Transport" );
			} else {
				jqXHR.readyState = 1;
	
				// Send global event
				if ( fireGlobals ) {
					globalEventContext.trigger( "ajaxSend", [ jqXHR, s ] );
				}
	
				// If request was aborted inside ajaxSend, stop there
				if ( completed ) {
					return jqXHR;
				}
	
				// Timeout
				if ( s.async && s.timeout > 0 ) {
					timeoutTimer = window.setTimeout( function() {
						jqXHR.abort( "timeout" );
					}, s.timeout );
				}
	
				try {
					completed = false;
					transport.send( requestHeaders, done );
				} catch ( e ) {
	
					// Rethrow post-completion exceptions
					if ( completed ) {
						throw e;
					}
	
					// Propagate others as results
					done( -1, e );
				}
			}
	
			// Callback for when everything is done
			function done( status, nativeStatusText, responses, headers ) {
				var isSuccess, success, error, response, modified,
					statusText = nativeStatusText;
	
				// Ignore repeat invocations
				if ( completed ) {
					return;
				}
	
				completed = true;
	
				// Clear timeout if it exists
				if ( timeoutTimer ) {
					window.clearTimeout( timeoutTimer );
				}
	
				// Dereference transport for early garbage collection
				// (no matter how long the jqXHR object will be used)
				transport = undefined;
	
				// Cache response headers
				responseHeadersString = headers || "";
	
				// Set readyState
				jqXHR.readyState = status > 0 ? 4 : 0;
	
				// Determine if successful
				isSuccess = status >= 200 && status < 300 || status === 304;
	
				// Get response data
				if ( responses ) {
					response = ajaxHandleResponses( s, jqXHR, responses );
				}
	
				// Convert no matter what (that way responseXXX fields are always set)
				response = ajaxConvert( s, response, jqXHR, isSuccess );
	
				// If successful, handle type chaining
				if ( isSuccess ) {
	
					// Set the If-Modified-Since and/or If-None-Match header, if in ifModified mode.
					if ( s.ifModified ) {
						modified = jqXHR.getResponseHeader( "Last-Modified" );
						if ( modified ) {
							jQuery.lastModified[ cacheURL ] = modified;
						}
						modified = jqXHR.getResponseHeader( "etag" );
						if ( modified ) {
							jQuery.etag[ cacheURL ] = modified;
						}
					}
	
					// if no content
					if ( status === 204 || s.type === "HEAD" ) {
						statusText = "nocontent";
	
					// if not modified
					} else if ( status === 304 ) {
						statusText = "notmodified";
	
					// If we have data, let's convert it
					} else {
						statusText = response.state;
						success = response.data;
						error = response.error;
						isSuccess = !error;
					}
				} else {
	
					// Extract error from statusText and normalize for non-aborts
					error = statusText;
					if ( status || !statusText ) {
						statusText = "error";
						if ( status < 0 ) {
							status = 0;
						}
					}
				}
	
				// Set data for the fake xhr object
				jqXHR.status = status;
				jqXHR.statusText = ( nativeStatusText || statusText ) + "";
	
				// Success/Error
				if ( isSuccess ) {
					deferred.resolveWith( callbackContext, [ success, statusText, jqXHR ] );
				} else {
					deferred.rejectWith( callbackContext, [ jqXHR, statusText, error ] );
				}
	
				// Status-dependent callbacks
				jqXHR.statusCode( statusCode );
				statusCode = undefined;
	
				if ( fireGlobals ) {
					globalEventContext.trigger( isSuccess ? "ajaxSuccess" : "ajaxError",
						[ jqXHR, s, isSuccess ? success : error ] );
				}
	
				// Complete
				completeDeferred.fireWith( callbackContext, [ jqXHR, statusText ] );
	
				if ( fireGlobals ) {
					globalEventContext.trigger( "ajaxComplete", [ jqXHR, s ] );
	
					// Handle the global AJAX counter
					if ( !( --jQuery.active ) ) {
						jQuery.event.trigger( "ajaxStop" );
					}
				}
			}
	
			return jqXHR;
		},
	
		getJSON: function( url, data, callback ) {
			return jQuery.get( url, data, callback, "json" );
		},
	
		getScript: function( url, callback ) {
			return jQuery.get( url, undefined, callback, "script" );
		}
	} );
	
	jQuery.each( [ "get", "post" ], function( i, method ) {
		jQuery[ method ] = function( url, data, callback, type ) {
	
			// Shift arguments if data argument was omitted
			if ( jQuery.isFunction( data ) ) {
				type = type || callback;
				callback = data;
				data = undefined;
			}
	
			// The url can be an options object (which then must have .url)
			return jQuery.ajax( jQuery.extend( {
				url: url,
				type: method,
				dataType: type,
				data: data,
				success: callback
			}, jQuery.isPlainObject( url ) && url ) );
		};
	} );
	
	
	jQuery._evalUrl = function( url ) {
		return jQuery.ajax( {
			url: url,
	
			// Make this explicit, since user can override this through ajaxSetup (#11264)
			type: "GET",
			dataType: "script",
			cache: true,
			async: false,
			global: false,
			"throws": true
		} );
	};
	
	
	jQuery.fn.extend( {
		wrapAll: function( html ) {
			var wrap;
	
			if ( this[ 0 ] ) {
				if ( jQuery.isFunction( html ) ) {
					html = html.call( this[ 0 ] );
				}
	
				// The elements to wrap the target around
				wrap = jQuery( html, this[ 0 ].ownerDocument ).eq( 0 ).clone( true );
	
				if ( this[ 0 ].parentNode ) {
					wrap.insertBefore( this[ 0 ] );
				}
	
				wrap.map( function() {
					var elem = this;
	
					while ( elem.firstElementChild ) {
						elem = elem.firstElementChild;
					}
	
					return elem;
				} ).append( this );
			}
	
			return this;
		},
	
		wrapInner: function( html ) {
			if ( jQuery.isFunction( html ) ) {
				return this.each( function( i ) {
					jQuery( this ).wrapInner( html.call( this, i ) );
				} );
			}
	
			return this.each( function() {
				var self = jQuery( this ),
					contents = self.contents();
	
				if ( contents.length ) {
					contents.wrapAll( html );
	
				} else {
					self.append( html );
				}
			} );
		},
	
		wrap: function( html ) {
			var isFunction = jQuery.isFunction( html );
	
			return this.each( function( i ) {
				jQuery( this ).wrapAll( isFunction ? html.call( this, i ) : html );
			} );
		},
	
		unwrap: function( selector ) {
			this.parent( selector ).not( "body" ).each( function() {
				jQuery( this ).replaceWith( this.childNodes );
			} );
			return this;
		}
	} );
	
	
	jQuery.expr.pseudos.hidden = function( elem ) {
		return !jQuery.expr.pseudos.visible( elem );
	};
	jQuery.expr.pseudos.visible = function( elem ) {
		return !!( elem.offsetWidth || elem.offsetHeight || elem.getClientRects().length );
	};
	
	
	
	
	jQuery.ajaxSettings.xhr = function() {
		try {
			return new window.XMLHttpRequest();
		} catch ( e ) {}
	};
	
	var xhrSuccessStatus = {
	
			// File protocol always yields status code 0, assume 200
			0: 200,
	
			// Support: IE <=9 only
			// #1450: sometimes IE returns 1223 when it should be 204
			1223: 204
		},
		xhrSupported = jQuery.ajaxSettings.xhr();
	
	support.cors = !!xhrSupported && ( "withCredentials" in xhrSupported );
	support.ajax = xhrSupported = !!xhrSupported;
	
	jQuery.ajaxTransport( function( options ) {
		var callback, errorCallback;
	
		// Cross domain only allowed if supported through XMLHttpRequest
		if ( support.cors || xhrSupported && !options.crossDomain ) {
			return {
				send: function( headers, complete ) {
					var i,
						xhr = options.xhr();
	
					xhr.open(
						options.type,
						options.url,
						options.async,
						options.username,
						options.password
					);
	
					// Apply custom fields if provided
					if ( options.xhrFields ) {
						for ( i in options.xhrFields ) {
							xhr[ i ] = options.xhrFields[ i ];
						}
					}
	
					// Override mime type if needed
					if ( options.mimeType && xhr.overrideMimeType ) {
						xhr.overrideMimeType( options.mimeType );
					}
	
					// X-Requested-With header
					// For cross-domain requests, seeing as conditions for a preflight are
					// akin to a jigsaw puzzle, we simply never set it to be sure.
					// (it can always be set on a per-request basis or even using ajaxSetup)
					// For same-domain requests, won't change header if already provided.
					if ( !options.crossDomain && !headers[ "X-Requested-With" ] ) {
						headers[ "X-Requested-With" ] = "XMLHttpRequest";
					}
	
					// Set headers
					for ( i in headers ) {
						xhr.setRequestHeader( i, headers[ i ] );
					}
	
					// Callback
					callback = function( type ) {
						return function() {
							if ( callback ) {
								callback = errorCallback = xhr.onload =
									xhr.onerror = xhr.onabort = xhr.onreadystatechange = null;
	
								if ( type === "abort" ) {
									xhr.abort();
								} else if ( type === "error" ) {
	
									// Support: IE <=9 only
									// On a manual native abort, IE9 throws
									// errors on any property access that is not readyState
									if ( typeof xhr.status !== "number" ) {
										complete( 0, "error" );
									} else {
										complete(
	
											// File: protocol always yields status 0; see #8605, #14207
											xhr.status,
											xhr.statusText
										);
									}
								} else {
									complete(
										xhrSuccessStatus[ xhr.status ] || xhr.status,
										xhr.statusText,
	
										// Support: IE <=9 only
										// IE9 has no XHR2 but throws on binary (trac-11426)
										// For XHR2 non-text, let the caller handle it (gh-2498)
										( xhr.responseType || "text" ) !== "text"  ||
										typeof xhr.responseText !== "string" ?
											{ binary: xhr.response } :
											{ text: xhr.responseText },
										xhr.getAllResponseHeaders()
									);
								}
							}
						};
					};
	
					// Listen to events
					xhr.onload = callback();
					errorCallback = xhr.onerror = callback( "error" );
	
					// Support: IE 9 only
					// Use onreadystatechange to replace onabort
					// to handle uncaught aborts
					if ( xhr.onabort !== undefined ) {
						xhr.onabort = errorCallback;
					} else {
						xhr.onreadystatechange = function() {
	
							// Check readyState before timeout as it changes
							if ( xhr.readyState === 4 ) {
	
								// Allow onerror to be called first,
								// but that will not handle a native abort
								// Also, save errorCallback to a variable
								// as xhr.onerror cannot be accessed
								window.setTimeout( function() {
									if ( callback ) {
										errorCallback();
									}
								} );
							}
						};
					}
	
					// Create the abort callback
					callback = callback( "abort" );
	
					try {
	
						// Do send the request (this may raise an exception)
						xhr.send( options.hasContent && options.data || null );
					} catch ( e ) {
	
						// #14683: Only rethrow if this hasn't been notified as an error yet
						if ( callback ) {
							throw e;
						}
					}
				},
	
				abort: function() {
					if ( callback ) {
						callback();
					}
				}
			};
		}
	} );
	
	
	
	
	// Prevent auto-execution of scripts when no explicit dataType was provided (See gh-2432)
	jQuery.ajaxPrefilter( function( s ) {
		if ( s.crossDomain ) {
			s.contents.script = false;
		}
	} );
	
	// Install script dataType
	jQuery.ajaxSetup( {
		accepts: {
			script: "text/javascript, application/javascript, " +
				"application/ecmascript, application/x-ecmascript"
		},
		contents: {
			script: /\b(?:java|ecma)script\b/
		},
		converters: {
			"text script": function( text ) {
				jQuery.globalEval( text );
				return text;
			}
		}
	} );
	
	// Handle cache's special case and crossDomain
	jQuery.ajaxPrefilter( "script", function( s ) {
		if ( s.cache === undefined ) {
			s.cache = false;
		}
		if ( s.crossDomain ) {
			s.type = "GET";
		}
	} );
	
	// Bind script tag hack transport
	jQuery.ajaxTransport( "script", function( s ) {
	
		// This transport only deals with cross domain requests
		if ( s.crossDomain ) {
			var script, callback;
			return {
				send: function( _, complete ) {
					script = jQuery( "<script>" ).prop( {
						charset: s.scriptCharset,
						src: s.url
					} ).on(
						"load error",
						callback = function( evt ) {
							script.remove();
							callback = null;
							if ( evt ) {
								complete( evt.type === "error" ? 404 : 200, evt.type );
							}
						}
					);
	
					// Use native DOM manipulation to avoid our domManip AJAX trickery
					document.head.appendChild( script[ 0 ] );
				},
				abort: function() {
					if ( callback ) {
						callback();
					}
				}
			};
		}
	} );
	
	
	
	
	var oldCallbacks = [],
		rjsonp = /(=)\?(?=&|$)|\?\?/;
	
	// Default jsonp settings
	jQuery.ajaxSetup( {
		jsonp: "callback",
		jsonpCallback: function() {
			var callback = oldCallbacks.pop() || ( jQuery.expando + "_" + ( nonce++ ) );
			this[ callback ] = true;
			return callback;
		}
	} );
	
	// Detect, normalize options and install callbacks for jsonp requests
	jQuery.ajaxPrefilter( "json jsonp", function( s, originalSettings, jqXHR ) {
	
		var callbackName, overwritten, responseContainer,
			jsonProp = s.jsonp !== false && ( rjsonp.test( s.url ) ?
				"url" :
				typeof s.data === "string" &&
					( s.contentType || "" )
						.indexOf( "application/x-www-form-urlencoded" ) === 0 &&
					rjsonp.test( s.data ) && "data"
			);
	
		// Handle iff the expected data type is "jsonp" or we have a parameter to set
		if ( jsonProp || s.dataTypes[ 0 ] === "jsonp" ) {
	
			// Get callback name, remembering preexisting value associated with it
			callbackName = s.jsonpCallback = jQuery.isFunction( s.jsonpCallback ) ?
				s.jsonpCallback() :
				s.jsonpCallback;
	
			// Insert callback into url or form data
			if ( jsonProp ) {
				s[ jsonProp ] = s[ jsonProp ].replace( rjsonp, "$1" + callbackName );
			} else if ( s.jsonp !== false ) {
				s.url += ( rquery.test( s.url ) ? "&" : "?" ) + s.jsonp + "=" + callbackName;
			}
	
			// Use data converter to retrieve json after script execution
			s.converters[ "script json" ] = function() {
				if ( !responseContainer ) {
					jQuery.error( callbackName + " was not called" );
				}
				return responseContainer[ 0 ];
			};
	
			// Force json dataType
			s.dataTypes[ 0 ] = "json";
	
			// Install callback
			overwritten = window[ callbackName ];
			window[ callbackName ] = function() {
				responseContainer = arguments;
			};
	
			// Clean-up function (fires after converters)
			jqXHR.always( function() {
	
				// If previous value didn't exist - remove it
				if ( overwritten === undefined ) {
					jQuery( window ).removeProp( callbackName );
	
				// Otherwise restore preexisting value
				} else {
					window[ callbackName ] = overwritten;
				}
	
				// Save back as free
				if ( s[ callbackName ] ) {
	
					// Make sure that re-using the options doesn't screw things around
					s.jsonpCallback = originalSettings.jsonpCallback;
	
					// Save the callback name for future use
					oldCallbacks.push( callbackName );
				}
	
				// Call if it was a function and we have a response
				if ( responseContainer && jQuery.isFunction( overwritten ) ) {
					overwritten( responseContainer[ 0 ] );
				}
	
				responseContainer = overwritten = undefined;
			} );
	
			// Delegate to script
			return "script";
		}
	} );
	
	
	
	
	// Support: Safari 8 only
	// In Safari 8 documents created via document.implementation.createHTMLDocument
	// collapse sibling forms: the second one becomes a child of the first one.
	// Because of that, this security measure has to be disabled in Safari 8.
	// https://bugs.webkit.org/show_bug.cgi?id=137337
	support.createHTMLDocument = ( function() {
		var body = document.implementation.createHTMLDocument( "" ).body;
		body.innerHTML = "<form></form><form></form>";
		return body.childNodes.length === 2;
	} )();
	
	
	// Argument "data" should be string of html
	// context (optional): If specified, the fragment will be created in this context,
	// defaults to document
	// keepScripts (optional): If true, will include scripts passed in the html string
	jQuery.parseHTML = function( data, context, keepScripts ) {
		if ( typeof data !== "string" ) {
			return [];
		}
		if ( typeof context === "boolean" ) {
			keepScripts = context;
			context = false;
		}
	
		var base, parsed, scripts;
	
		if ( !context ) {
	
			// Stop scripts or inline event handlers from being executed immediately
			// by using document.implementation
			if ( support.createHTMLDocument ) {
				context = document.implementation.createHTMLDocument( "" );
	
				// Set the base href for the created document
				// so any parsed elements with URLs
				// are based on the document's URL (gh-2965)
				base = context.createElement( "base" );
				base.href = document.location.href;
				context.head.appendChild( base );
			} else {
				context = document;
			}
		}
	
		parsed = rsingleTag.exec( data );
		scripts = !keepScripts && [];
	
		// Single tag
		if ( parsed ) {
			return [ context.createElement( parsed[ 1 ] ) ];
		}
	
		parsed = buildFragment( [ data ], context, scripts );
	
		if ( scripts && scripts.length ) {
			jQuery( scripts ).remove();
		}
	
		return jQuery.merge( [], parsed.childNodes );
	};
	
	
	/**
	 * Load a url into a page
	 */
	jQuery.fn.load = function( url, params, callback ) {
		var selector, type, response,
			self = this,
			off = url.indexOf( " " );
	
		if ( off > -1 ) {
			selector = stripAndCollapse( url.slice( off ) );
			url = url.slice( 0, off );
		}
	
		// If it's a function
		if ( jQuery.isFunction( params ) ) {
	
			// We assume that it's the callback
			callback = params;
			params = undefined;
	
		// Otherwise, build a param string
		} else if ( params && typeof params === "object" ) {
			type = "POST";
		}
	
		// If we have elements to modify, make the request
		if ( self.length > 0 ) {
			jQuery.ajax( {
				url: url,
	
				// If "type" variable is undefined, then "GET" method will be used.
				// Make value of this field explicit since
				// user can override it through ajaxSetup method
				type: type || "GET",
				dataType: "html",
				data: params
			} ).done( function( responseText ) {
	
				// Save response for use in complete callback
				response = arguments;
	
				self.html( selector ?
	
					// If a selector was specified, locate the right elements in a dummy div
					// Exclude scripts to avoid IE 'Permission Denied' errors
					jQuery( "<div>" ).append( jQuery.parseHTML( responseText ) ).find( selector ) :
	
					// Otherwise use the full result
					responseText );
	
			// If the request succeeds, this function gets "data", "status", "jqXHR"
			// but they are ignored because response was set above.
			// If it fails, this function gets "jqXHR", "status", "error"
			} ).always( callback && function( jqXHR, status ) {
				self.each( function() {
					callback.apply( this, response || [ jqXHR.responseText, status, jqXHR ] );
				} );
			} );
		}
	
		return this;
	};
	
	
	
	
	// Attach a bunch of functions for handling common AJAX events
	jQuery.each( [
		"ajaxStart",
		"ajaxStop",
		"ajaxComplete",
		"ajaxError",
		"ajaxSuccess",
		"ajaxSend"
	], function( i, type ) {
		jQuery.fn[ type ] = function( fn ) {
			return this.on( type, fn );
		};
	} );
	
	
	
	
	jQuery.expr.pseudos.animated = function( elem ) {
		return jQuery.grep( jQuery.timers, function( fn ) {
			return elem === fn.elem;
		} ).length;
	};
	
	
	
	
	jQuery.offset = {
		setOffset: function( elem, options, i ) {
			var curPosition, curLeft, curCSSTop, curTop, curOffset, curCSSLeft, calculatePosition,
				position = jQuery.css( elem, "position" ),
				curElem = jQuery( elem ),
				props = {};
	
			// Set position first, in-case top/left are set even on static elem
			if ( position === "static" ) {
				elem.style.position = "relative";
			}
	
			curOffset = curElem.offset();
			curCSSTop = jQuery.css( elem, "top" );
			curCSSLeft = jQuery.css( elem, "left" );
			calculatePosition = ( position === "absolute" || position === "fixed" ) &&
				( curCSSTop + curCSSLeft ).indexOf( "auto" ) > -1;
	
			// Need to be able to calculate position if either
			// top or left is auto and position is either absolute or fixed
			if ( calculatePosition ) {
				curPosition = curElem.position();
				curTop = curPosition.top;
				curLeft = curPosition.left;
	
			} else {
				curTop = parseFloat( curCSSTop ) || 0;
				curLeft = parseFloat( curCSSLeft ) || 0;
			}
	
			if ( jQuery.isFunction( options ) ) {
	
				// Use jQuery.extend here to allow modification of coordinates argument (gh-1848)
				options = options.call( elem, i, jQuery.extend( {}, curOffset ) );
			}
	
			if ( options.top != null ) {
				props.top = ( options.top - curOffset.top ) + curTop;
			}
			if ( options.left != null ) {
				props.left = ( options.left - curOffset.left ) + curLeft;
			}
	
			if ( "using" in options ) {
				options.using.call( elem, props );
	
			} else {
				curElem.css( props );
			}
		}
	};
	
	jQuery.fn.extend( {
		offset: function( options ) {
	
			// Preserve chaining for setter
			if ( arguments.length ) {
				return options === undefined ?
					this :
					this.each( function( i ) {
						jQuery.offset.setOffset( this, options, i );
					} );
			}
	
			var doc, docElem, rect, win,
				elem = this[ 0 ];
	
			if ( !elem ) {
				return;
			}
	
			// Return zeros for disconnected and hidden (display: none) elements (gh-2310)
			// Support: IE <=11 only
			// Running getBoundingClientRect on a
			// disconnected node in IE throws an error
			if ( !elem.getClientRects().length ) {
				return { top: 0, left: 0 };
			}
	
			rect = elem.getBoundingClientRect();
	
			doc = elem.ownerDocument;
			docElem = doc.documentElement;
			win = doc.defaultView;
	
			return {
				top: rect.top + win.pageYOffset - docElem.clientTop,
				left: rect.left + win.pageXOffset - docElem.clientLeft
			};
		},
	
		position: function() {
			if ( !this[ 0 ] ) {
				return;
			}
	
			var offsetParent, offset,
				elem = this[ 0 ],
				parentOffset = { top: 0, left: 0 };
	
			// Fixed elements are offset from window (parentOffset = {top:0, left: 0},
			// because it is its only offset parent
			if ( jQuery.css( elem, "position" ) === "fixed" ) {
	
				// Assume getBoundingClientRect is there when computed position is fixed
				offset = elem.getBoundingClientRect();
	
			} else {
	
				// Get *real* offsetParent
				offsetParent = this.offsetParent();
	
				// Get correct offsets
				offset = this.offset();
				if ( !nodeName( offsetParent[ 0 ], "html" ) ) {
					parentOffset = offsetParent.offset();
				}
	
				// Add offsetParent borders
				parentOffset = {
					top: parentOffset.top + jQuery.css( offsetParent[ 0 ], "borderTopWidth", true ),
					left: parentOffset.left + jQuery.css( offsetParent[ 0 ], "borderLeftWidth", true )
				};
			}
	
			// Subtract parent offsets and element margins
			return {
				top: offset.top - parentOffset.top - jQuery.css( elem, "marginTop", true ),
				left: offset.left - parentOffset.left - jQuery.css( elem, "marginLeft", true )
			};
		},
	
		// This method will return documentElement in the following cases:
		// 1) For the element inside the iframe without offsetParent, this method will return
		//    documentElement of the parent window
		// 2) For the hidden or detached element
		// 3) For body or html element, i.e. in case of the html node - it will return itself
		//
		// but those exceptions were never presented as a real life use-cases
		// and might be considered as more preferable results.
		//
		// This logic, however, is not guaranteed and can change at any point in the future
		offsetParent: function() {
			return this.map( function() {
				var offsetParent = this.offsetParent;
	
				while ( offsetParent && jQuery.css( offsetParent, "position" ) === "static" ) {
					offsetParent = offsetParent.offsetParent;
				}
	
				return offsetParent || documentElement;
			} );
		}
	} );
	
	// Create scrollLeft and scrollTop methods
	jQuery.each( { scrollLeft: "pageXOffset", scrollTop: "pageYOffset" }, function( method, prop ) {
		var top = "pageYOffset" === prop;
	
		jQuery.fn[ method ] = function( val ) {
			return access( this, function( elem, method, val ) {
	
				// Coalesce documents and windows
				var win;
				if ( jQuery.isWindow( elem ) ) {
					win = elem;
				} else if ( elem.nodeType === 9 ) {
					win = elem.defaultView;
				}
	
				if ( val === undefined ) {
					return win ? win[ prop ] : elem[ method ];
				}
	
				if ( win ) {
					win.scrollTo(
						!top ? val : win.pageXOffset,
						top ? val : win.pageYOffset
					);
	
				} else {
					elem[ method ] = val;
				}
			}, method, val, arguments.length );
		};
	} );
	
	// Support: Safari <=7 - 9.1, Chrome <=37 - 49
	// Add the top/left cssHooks using jQuery.fn.position
	// Webkit bug: https://bugs.webkit.org/show_bug.cgi?id=29084
	// Blink bug: https://bugs.chromium.org/p/chromium/issues/detail?id=589347
	// getComputedStyle returns percent when specified for top/left/bottom/right;
	// rather than make the css module depend on the offset module, just check for it here
	jQuery.each( [ "top", "left" ], function( i, prop ) {
		jQuery.cssHooks[ prop ] = addGetHookIf( support.pixelPosition,
			function( elem, computed ) {
				if ( computed ) {
					computed = curCSS( elem, prop );
	
					// If curCSS returns percentage, fallback to offset
					return rnumnonpx.test( computed ) ?
						jQuery( elem ).position()[ prop ] + "px" :
						computed;
				}
			}
		);
	} );
	
	
	// Create innerHeight, innerWidth, height, width, outerHeight and outerWidth methods
	jQuery.each( { Height: "height", Width: "width" }, function( name, type ) {
		jQuery.each( { padding: "inner" + name, content: type, "": "outer" + name },
			function( defaultExtra, funcName ) {
	
			// Margin is only for outerHeight, outerWidth
			jQuery.fn[ funcName ] = function( margin, value ) {
				var chainable = arguments.length && ( defaultExtra || typeof margin !== "boolean" ),
					extra = defaultExtra || ( margin === true || value === true ? "margin" : "border" );
	
				return access( this, function( elem, type, value ) {
					var doc;
	
					if ( jQuery.isWindow( elem ) ) {
	
						// $( window ).outerWidth/Height return w/h including scrollbars (gh-1729)
						return funcName.indexOf( "outer" ) === 0 ?
							elem[ "inner" + name ] :
							elem.document.documentElement[ "client" + name ];
					}
	
					// Get document width or height
					if ( elem.nodeType === 9 ) {
						doc = elem.documentElement;
	
						// Either scroll[Width/Height] or offset[Width/Height] or client[Width/Height],
						// whichever is greatest
						return Math.max(
							elem.body[ "scroll" + name ], doc[ "scroll" + name ],
							elem.body[ "offset" + name ], doc[ "offset" + name ],
							doc[ "client" + name ]
						);
					}
	
					return value === undefined ?
	
						// Get width or height on the element, requesting but not forcing parseFloat
						jQuery.css( elem, type, extra ) :
	
						// Set width or height on the element
						jQuery.style( elem, type, value, extra );
				}, type, chainable ? margin : undefined, chainable );
			};
		} );
	} );
	
	
	jQuery.fn.extend( {
	
		bind: function( types, data, fn ) {
			return this.on( types, null, data, fn );
		},
		unbind: function( types, fn ) {
			return this.off( types, null, fn );
		},
	
		delegate: function( selector, types, data, fn ) {
			return this.on( types, selector, data, fn );
		},
		undelegate: function( selector, types, fn ) {
	
			// ( namespace ) or ( selector, types [, fn] )
			return arguments.length === 1 ?
				this.off( selector, "**" ) :
				this.off( types, selector || "**", fn );
		}
	} );
	
	jQuery.holdReady = function( hold ) {
		if ( hold ) {
			jQuery.readyWait++;
		} else {
			jQuery.ready( true );
		}
	};
	jQuery.isArray = Array.isArray;
	jQuery.parseJSON = JSON.parse;
	jQuery.nodeName = nodeName;
	
	
	
	
	// Register as a named AMD module, since jQuery can be concatenated with other
	// files that may use define, but not via a proper concatenation script that
	// understands anonymous AMD modules. A named AMD is safest and most robust
	// way to register. Lowercase jquery is used because AMD module names are
	// derived from file names, and jQuery is normally delivered in a lowercase
	// file name. Do this after creating the global so that if an AMD module wants
	// to call noConflict to hide this version of jQuery, it will work.
	
	// Note that for maximum portability, libraries that are not jQuery should
	// declare themselves as anonymous modules, and avoid setting a global if an
	// AMD loader is present. jQuery is a special case. For more information, see
	// https://github.com/jrburke/requirejs/wiki/Updating-existing-libraries#wiki-anon
	
	if ( true ) {
		!(__WEBPACK_AMD_DEFINE_ARRAY__ = [], __WEBPACK_AMD_DEFINE_RESULT__ = function() {
			return jQuery;
		}.apply(exports, __WEBPACK_AMD_DEFINE_ARRAY__), __WEBPACK_AMD_DEFINE_RESULT__ !== undefined && (module.exports = __WEBPACK_AMD_DEFINE_RESULT__));
	}
	
	
	
	
	var
	
		// Map over jQuery in case of overwrite
		_jQuery = window.jQuery,
	
		// Map over the $ in case of overwrite
		_$ = window.$;
	
	jQuery.noConflict = function( deep ) {
		if ( window.$ === jQuery ) {
			window.$ = _$;
		}
	
		if ( deep && window.jQuery === jQuery ) {
			window.jQuery = _jQuery;
		}
	
		return jQuery;
	};
	
	// Expose jQuery and $ identifiers, even in AMD
	// (#7102#comment:10, https://github.com/jquery/jquery/pull/557)
	// and CommonJS for browser emulators (#13566)
	if ( !noGlobal ) {
		window.jQuery = window.$ = jQuery;
	}
	
	
	
	
	return jQuery;
	} );


/***/ },
/* 11 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	function __export(m) {
	    for (var p in m) if (!exports.hasOwnProperty(p)) exports[p] = m[p];
	}
	Object.defineProperty(exports, "__esModule", { value: true });
	var util_1 = __webpack_require__(12);
	var storage_1 = __webpack_require__(17);
	var configuration_1 = __webpack_require__(19);
	var logging_1 = __webpack_require__(14);
	var rest_1 = __webpack_require__(20);
	var pnplibconfig_1 = __webpack_require__(15);
	/**
	 * Root class of the Patterns and Practices namespace, provides an entry point to the library
	 */
	/**
	 * Utility methods
	 */
	exports.util = util_1.Util;
	/**
	 * Provides access to the REST interface
	 */
	exports.sp = new rest_1.Rest();
	/**
	 * Provides access to local and session storage
	 */
	exports.storage = new storage_1.PnPClientStorage();
	/**
	 * Global configuration instance to which providers can be added
	 */
	exports.config = new configuration_1.Settings();
	/**
	 * Global logging instance to which subscribers can be registered and messages written
	 */
	exports.log = logging_1.Logger;
	/**
	 * Allows for the configuration of the library
	 */
	exports.setup = pnplibconfig_1.setRuntimeConfig;
	/**
	 * Expose a subset of classes from the library for public consumption
	 */
	__export(__webpack_require__(55));
	// creating this class instead of directly assigning to default fixes issue #116
	var Def = {
	    /**
	     * Global configuration instance to which providers can be added
	     */
	    config: exports.config,
	    /**
	     * Global logging instance to which subscribers can be registered and messages written
	     */
	    log: exports.log,
	    /**
	     * Provides access to local and session storage
	     */
	    setup: exports.setup,
	    /**
	     * Provides access to the REST interface
	     */
	    sp: exports.sp,
	    /**
	     * Provides access to local and session storage
	     */
	    storage: exports.storage,
	    /**
	     * Utility methods
	     */
	    util: exports.util,
	};
	/**
	 * Enables use of the import pnp from syntax
	 */
	exports.default = Def;


/***/ },
/* 12 */
/***/ function(module, exports, __webpack_require__) {

	/* WEBPACK VAR INJECTION */(function(global) {"use strict";
	var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
	    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
	    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
	    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
	    return c > 3 && r && Object.defineProperty(target, key, r), r;
	};
	Object.defineProperty(exports, "__esModule", { value: true });
	var decorators_1 = __webpack_require__(13);
	var pnplibconfig_1 = __webpack_require__(15);
	var Util = (function () {
	    function Util() {
	    }
	    /**
	     * Gets a callback function which will maintain context across async calls.
	     * Allows for the calling pattern getCtxCallback(thisobj, method, methodarg1, methodarg2, ...)
	     *
	     * @param context The object that will be the 'this' value in the callback
	     * @param method The method to which we will apply the context and parameters
	     * @param params Optional, additional arguments to supply to the wrapped method when it is invoked
	     */
	    Util.getCtxCallback = function (context, method) {
	        var params = [];
	        for (var _i = 2; _i < arguments.length; _i++) {
	            params[_i - 2] = arguments[_i];
	        }
	        return function () {
	            method.apply(context, params);
	        };
	    };
	    /**
	     * Tests if a url param exists
	     *
	     * @param name The name of the url paramter to check
	     */
	    Util.urlParamExists = function (name) {
	        name = name.replace(/[\[]/, "\\[").replace(/[\]]/, "\\]");
	        var regex = new RegExp("[\\?&]" + name + "=([^&#]*)");
	        return regex.test(location.search);
	    };
	    /**
	     * Gets a url param value by name
	     *
	     * @param name The name of the paramter for which we want the value
	     */
	    Util.getUrlParamByName = function (name) {
	        name = name.replace(/[\[]/, "\\[").replace(/[\]]/, "\\]");
	        var regex = new RegExp("[\\?&]" + name + "=([^&#]*)");
	        var results = regex.exec(location.search);
	        return results == null ? "" : decodeURIComponent(results[1].replace(/\+/g, " "));
	    };
	    /**
	     * Gets a url param by name and attempts to parse a bool value
	     *
	     * @param name The name of the paramter for which we want the boolean value
	     */
	    Util.getUrlParamBoolByName = function (name) {
	        var p = this.getUrlParamByName(name);
	        var isFalse = (p === "" || /false|0/i.test(p));
	        return !isFalse;
	    };
	    /**
	     * Inserts the string s into the string target as the index specified by index
	     *
	     * @param target The string into which we will insert s
	     * @param index The location in target to insert s (zero based)
	     * @param s The string to insert into target at position index
	     */
	    Util.stringInsert = function (target, index, s) {
	        if (index > 0) {
	            return target.substring(0, index) + s + target.substring(index, target.length);
	        }
	        return s + target;
	    };
	    /**
	     * Adds a value to a date
	     *
	     * @param date The date to which we will add units, done in local time
	     * @param interval The name of the interval to add, one of: ['year', 'quarter', 'month', 'week', 'day', 'hour', 'minute', 'second']
	     * @param units The amount to add to date of the given interval
	     *
	     * http://stackoverflow.com/questions/1197928/how-to-add-30-minutes-to-a-javascript-date-object
	     */
	    Util.dateAdd = function (date, interval, units) {
	        var ret = new Date(date.toLocaleString()); // don't change original date
	        switch (interval.toLowerCase()) {
	            case "year":
	                ret.setFullYear(ret.getFullYear() + units);
	                break;
	            case "quarter":
	                ret.setMonth(ret.getMonth() + 3 * units);
	                break;
	            case "month":
	                ret.setMonth(ret.getMonth() + units);
	                break;
	            case "week":
	                ret.setDate(ret.getDate() + 7 * units);
	                break;
	            case "day":
	                ret.setDate(ret.getDate() + units);
	                break;
	            case "hour":
	                ret.setTime(ret.getTime() + units * 3600000);
	                break;
	            case "minute":
	                ret.setTime(ret.getTime() + units * 60000);
	                break;
	            case "second":
	                ret.setTime(ret.getTime() + units * 1000);
	                break;
	            default:
	                ret = undefined;
	                break;
	        }
	        return ret;
	    };
	    /**
	     * Loads a stylesheet into the current page
	     *
	     * @param path The url to the stylesheet
	     * @param avoidCache If true a value will be appended as a query string to avoid browser caching issues
	     */
	    Util.loadStylesheet = function (path, avoidCache) {
	        if (avoidCache) {
	            path += "?" + encodeURIComponent((new Date()).getTime().toString());
	        }
	        var head = document.getElementsByTagName("head");
	        if (head.length > 0) {
	            var e = document.createElement("link");
	            head[0].appendChild(e);
	            e.setAttribute("type", "text/css");
	            e.setAttribute("rel", "stylesheet");
	            e.setAttribute("href", path);
	        }
	    };
	    /**
	     * Combines an arbitrary set of paths ensuring that the slashes are normalized
	     *
	     * @param paths 0 to n path parts to combine
	     */
	    Util.combinePaths = function () {
	        var paths = [];
	        for (var _i = 0; _i < arguments.length; _i++) {
	            paths[_i] = arguments[_i];
	        }
	        return paths
	            .filter(function (path) { return typeof path !== "undefined" && path !== null; })
	            .map(function (path) { return path.replace(/^[\\|\/]/, "").replace(/[\\|\/]$/, ""); })
	            .join("/")
	            .replace(/\\/g, "/");
	    };
	    /**
	     * Gets a random string of chars length
	     *
	     * @param chars The length of the random string to generate
	     */
	    Util.getRandomString = function (chars) {
	        var text = new Array(chars);
	        var possible = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
	        for (var i = 0; i < chars; i++) {
	            text[i] = possible.charAt(Math.floor(Math.random() * possible.length));
	        }
	        return text.join("");
	    };
	    /**
	     * Gets a random GUID value
	     *
	     * http://stackoverflow.com/questions/105034/create-guid-uuid-in-javascript
	     */
	    /* tslint:disable no-bitwise */
	    Util.getGUID = function () {
	        var d = new Date().getTime();
	        var guid = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, function (c) {
	            var r = (d + Math.random() * 16) % 16 | 0;
	            d = Math.floor(d / 16);
	            return (c === "x" ? r : (r & 0x3 | 0x8)).toString(16);
	        });
	        return guid;
	    };
	    /* tslint:enable */
	    /**
	     * Determines if a given value is a function
	     *
	     * @param candidateFunction The thing to test for being a function
	     */
	    Util.isFunction = function (candidateFunction) {
	        return typeof candidateFunction === "function";
	    };
	    /**
	     * @returns whether the provided parameter is a JavaScript Array or not.
	    */
	    Util.isArray = function (array) {
	        if (Array.isArray) {
	            return Array.isArray(array);
	        }
	        return array && typeof array.length === "number" && array.constructor === Array;
	    };
	    /**
	     * Determines if a string is null or empty or undefined
	     *
	     * @param s The string to test
	     */
	    Util.stringIsNullOrEmpty = function (s) {
	        return typeof s === "undefined" || s === null || s === "";
	    };
	    /**
	     * Provides functionality to extend the given object by doing a shallow copy
	     *
	     * @param target The object to which properties will be copied
	     * @param source The source object from which properties will be copied
	     * @param noOverwrite If true existing properties on the target are not overwritten from the source
	     *
	     */
	    Util.extend = function (target, source, noOverwrite) {
	        if (noOverwrite === void 0) { noOverwrite = false; }
	        if (source === null || typeof source === "undefined") {
	            return target;
	        }
	        // ensure we don't overwrite things we don't want overwritten
	        var check = noOverwrite ? function (o, i) { return !(i in o); } : function () { return true; };
	        return Object.getOwnPropertyNames(source)
	            .filter(function (v) { return check(target, v); })
	            .reduce(function (t, v) {
	            t[v] = source[v];
	            return t;
	        }, target);
	    };
	    /**
	     * Determines if a given url is absolute
	     *
	     * @param url The url to check to see if it is absolute
	     */
	    Util.isUrlAbsolute = function (url) {
	        return /^https?:\/\/|^\/\//i.test(url);
	    };
	    /**
	     * Attempts to make the supplied relative url absolute based on the _spPageContextInfo object, if available
	     *
	     * @param url The relative url to make absolute
	     */
	    Util.makeUrlAbsolute = function (url) {
	        if (Util.isUrlAbsolute(url)) {
	            return url;
	        }
	        if (typeof global._spPageContextInfo !== "undefined") {
	            if (global._spPageContextInfo.hasOwnProperty("webAbsoluteUrl")) {
	                return Util.combinePaths(global._spPageContextInfo.webAbsoluteUrl, url);
	            }
	            else if (global._spPageContextInfo.hasOwnProperty("webServerRelativeUrl")) {
	                return Util.combinePaths(global._spPageContextInfo.webServerRelativeUrl, url);
	            }
	        }
	        else {
	            return url;
	        }
	    };
	    /**
	     * Ensures that a given url is absolute for the current web based on context
	     *
	     * @param candidateUrl The url to make absolute
	     *
	     */
	    Util.toAbsoluteUrl = function (candidateUrl) {
	        return new Promise(function (resolve) {
	            if (Util.isUrlAbsolute(candidateUrl)) {
	                // if we are already absolute, then just return the url
	                return resolve(candidateUrl);
	            }
	            if (pnplibconfig_1.RuntimeConfig.baseUrl !== null) {
	                // base url specified either with baseUrl of spfxContext config property
	                return resolve(Util.combinePaths(pnplibconfig_1.RuntimeConfig.baseUrl, candidateUrl));
	            }
	            if (typeof global._spPageContextInfo !== "undefined") {
	                // operating in classic pages
	                if (global._spPageContextInfo.hasOwnProperty("webAbsoluteUrl")) {
	                    return resolve(Util.combinePaths(global._spPageContextInfo.webAbsoluteUrl, candidateUrl));
	                }
	                else if (global._spPageContextInfo.hasOwnProperty("webServerRelativeUrl")) {
	                    return resolve(Util.combinePaths(global._spPageContextInfo.webServerRelativeUrl, candidateUrl));
	                }
	            }
	            // does window.location exist and have _layouts in it?
	            if (typeof global.location !== "undefined") {
	                var index = global.location.toString().toLowerCase().indexOf("/_layouts/");
	                if (index > 0) {
	                    // we are likely in the workbench in /_layouts/
	                    return resolve(Util.combinePaths(global.location.toString().substr(0, index), candidateUrl));
	                }
	            }
	            return resolve(candidateUrl);
	        });
	    };
	    return Util;
	}());
	__decorate([
	    decorators_1.deprecated("The Util.makeUrlAbsolute method is deprecated and will be removed from future releases. Use Util.toAbsoluteUrl instead")
	], Util, "makeUrlAbsolute", null);
	exports.Util = Util;
	
	/* WEBPACK VAR INJECTION */}.call(exports, (function() { return this; }())))

/***/ },
/* 13 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	Object.defineProperty(exports, "__esModule", { value: true });
	var logging_1 = __webpack_require__(14);
	function deprecated(message) {
	    return function (target, propertyKey, descriptor) {
	        var method = descriptor.value;
	        descriptor.value = function () {
	            var args = [];
	            for (var _i = 0; _i < arguments.length; _i++) {
	                args[_i] = arguments[_i];
	            }
	            logging_1.Logger.log({
	                data: {
	                    descriptor: descriptor,
	                    propertyKey: propertyKey,
	                    target: target,
	                },
	                level: logging_1.LogLevel.Warning,
	                message: message,
	            });
	            return method.apply(this, args);
	        };
	    };
	}
	exports.deprecated = deprecated;


/***/ },
/* 14 */
/***/ function(module, exports) {

	"use strict";
	Object.defineProperty(exports, "__esModule", { value: true });
	/**
	 * A set of logging levels
	 *
	 */
	var LogLevel;
	(function (LogLevel) {
	    LogLevel[LogLevel["Verbose"] = 0] = "Verbose";
	    LogLevel[LogLevel["Info"] = 1] = "Info";
	    LogLevel[LogLevel["Warning"] = 2] = "Warning";
	    LogLevel[LogLevel["Error"] = 3] = "Error";
	    LogLevel[LogLevel["Off"] = 99] = "Off";
	})(LogLevel = exports.LogLevel || (exports.LogLevel = {}));
	/**
	 * Class used to subscribe ILogListener and log messages throughout an application
	 *
	 */
	var Logger = (function () {
	    function Logger() {
	    }
	    Object.defineProperty(Logger, "activeLogLevel", {
	        get: function () {
	            return Logger.instance.activeLogLevel;
	        },
	        set: function (value) {
	            Logger.instance.activeLogLevel = value;
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Logger, "instance", {
	        get: function () {
	            if (typeof Logger._instance === "undefined" || Logger._instance === null) {
	                Logger._instance = new LoggerImpl();
	            }
	            return Logger._instance;
	        },
	        enumerable: true,
	        configurable: true
	    });
	    /**
	     * Adds ILogListener instances to the set of subscribed listeners
	     *
	     * @param listeners One or more listeners to subscribe to this log
	     */
	    Logger.subscribe = function () {
	        var listeners = [];
	        for (var _i = 0; _i < arguments.length; _i++) {
	            listeners[_i] = arguments[_i];
	        }
	        listeners.map(function (listener) { return Logger.instance.subscribe(listener); });
	    };
	    /**
	     * Clears the subscribers collection, returning the collection before modifiction
	     */
	    Logger.clearSubscribers = function () {
	        return Logger.instance.clearSubscribers();
	    };
	    Object.defineProperty(Logger, "count", {
	        /**
	         * Gets the current subscriber count
	         */
	        get: function () {
	            return Logger.instance.count;
	        },
	        enumerable: true,
	        configurable: true
	    });
	    /**
	     * Writes the supplied string to the subscribed listeners
	     *
	     * @param message The message to write
	     * @param level [Optional] if supplied will be used as the level of the entry (Default: LogLevel.Verbose)
	     */
	    Logger.write = function (message, level) {
	        if (level === void 0) { level = LogLevel.Verbose; }
	        Logger.instance.log({ level: level, message: message });
	    };
	    /**
	     * Writes the supplied string to the subscribed listeners
	     *
	     * @param json The json object to stringify and write
	     * @param level [Optional] if supplied will be used as the level of the entry (Default: LogLevel.Verbose)
	     */
	    Logger.writeJSON = function (json, level) {
	        if (level === void 0) { level = LogLevel.Verbose; }
	        Logger.instance.log({ level: level, message: JSON.stringify(json) });
	    };
	    /**
	     * Logs the supplied entry to the subscribed listeners
	     *
	     * @param entry The message to log
	     */
	    Logger.log = function (entry) {
	        Logger.instance.log(entry);
	    };
	    /**
	     * Logs performance tracking data for the the execution duration of the supplied function using console.profile
	     *
	     * @param name The name of this profile boundary
	     * @param f The function to execute and track within this performance boundary
	     */
	    Logger.measure = function (name, f) {
	        return Logger.instance.measure(name, f);
	    };
	    return Logger;
	}());
	exports.Logger = Logger;
	var LoggerImpl = (function () {
	    function LoggerImpl(activeLogLevel, subscribers) {
	        if (activeLogLevel === void 0) { activeLogLevel = LogLevel.Warning; }
	        if (subscribers === void 0) { subscribers = []; }
	        this.activeLogLevel = activeLogLevel;
	        this.subscribers = subscribers;
	    }
	    LoggerImpl.prototype.subscribe = function (listener) {
	        this.subscribers.push(listener);
	    };
	    LoggerImpl.prototype.clearSubscribers = function () {
	        var s = this.subscribers.slice(0);
	        this.subscribers.length = 0;
	        return s;
	    };
	    Object.defineProperty(LoggerImpl.prototype, "count", {
	        get: function () {
	            return this.subscribers.length;
	        },
	        enumerable: true,
	        configurable: true
	    });
	    LoggerImpl.prototype.write = function (message, level) {
	        if (level === void 0) { level = LogLevel.Verbose; }
	        this.log({ level: level, message: message });
	    };
	    LoggerImpl.prototype.log = function (entry) {
	        if (typeof entry === "undefined" || entry.level < this.activeLogLevel) {
	            return;
	        }
	        this.subscribers.map(function (subscriber) { return subscriber.log(entry); });
	    };
	    LoggerImpl.prototype.measure = function (name, f) {
	        console.profile(name);
	        try {
	            return f();
	        }
	        finally {
	            console.profileEnd();
	        }
	    };
	    return LoggerImpl;
	}());
	/**
	 * Implementation of ILogListener which logs to the browser console
	 *
	 */
	var ConsoleListener = (function () {
	    function ConsoleListener() {
	    }
	    /**
	     * Any associated data that a given logging listener may choose to log or ignore
	     *
	     * @param entry The information to be logged
	     */
	    ConsoleListener.prototype.log = function (entry) {
	        var msg = this.format(entry);
	        switch (entry.level) {
	            case LogLevel.Verbose:
	            case LogLevel.Info:
	                console.log(msg);
	                break;
	            case LogLevel.Warning:
	                console.warn(msg);
	                break;
	            case LogLevel.Error:
	                console.error(msg);
	                break;
	        }
	    };
	    /**
	     * Formats the message
	     *
	     * @param entry The information to format into a string
	     */
	    ConsoleListener.prototype.format = function (entry) {
	        return "Message: " + entry.message + " Data: " + JSON.stringify(entry.data);
	    };
	    return ConsoleListener;
	}());
	exports.ConsoleListener = ConsoleListener;
	/**
	 * Implementation of ILogListener which logs to the supplied function
	 *
	 */
	var FunctionListener = (function () {
	    /**
	     * Creates a new instance of the FunctionListener class
	     *
	     * @constructor
	     * @param  method The method to which any logging data will be passed
	     */
	    function FunctionListener(method) {
	        this.method = method;
	    }
	    /**
	     * Any associated data that a given logging listener may choose to log or ignore
	     *
	     * @param entry The information to be logged
	     */
	    FunctionListener.prototype.log = function (entry) {
	        this.method(entry);
	    };
	    return FunctionListener;
	}());
	exports.FunctionListener = FunctionListener;


/***/ },
/* 15 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	Object.defineProperty(exports, "__esModule", { value: true });
	var fetchclient_1 = __webpack_require__(16);
	var RuntimeConfigImpl = (function () {
	    function RuntimeConfigImpl() {
	        // these are our default values for the library
	        this._headers = null;
	        this._defaultCachingStore = "session";
	        this._defaultCachingTimeoutSeconds = 60;
	        this._globalCacheDisable = false;
	        this._fetchClientFactory = function () { return new fetchclient_1.FetchClient(); };
	        this._baseUrl = null;
	        this._spfxContext = null;
	    }
	    RuntimeConfigImpl.prototype.set = function (config) {
	        if (config.hasOwnProperty("headers")) {
	            this._headers = config.headers;
	        }
	        if (config.hasOwnProperty("globalCacheDisable")) {
	            this._globalCacheDisable = config.globalCacheDisable;
	        }
	        if (config.hasOwnProperty("defaultCachingStore")) {
	            this._defaultCachingStore = config.defaultCachingStore;
	        }
	        if (config.hasOwnProperty("defaultCachingTimeoutSeconds")) {
	            this._defaultCachingTimeoutSeconds = config.defaultCachingTimeoutSeconds;
	        }
	        if (config.hasOwnProperty("fetchClientFactory")) {
	            this._fetchClientFactory = config.fetchClientFactory;
	        }
	        if (config.hasOwnProperty("baseUrl")) {
	            this._baseUrl = config.baseUrl;
	        }
	        if (config.hasOwnProperty("spfxContext")) {
	            this._spfxContext = config.spfxContext;
	        }
	    };
	    Object.defineProperty(RuntimeConfigImpl.prototype, "headers", {
	        get: function () {
	            return this._headers;
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(RuntimeConfigImpl.prototype, "defaultCachingStore", {
	        get: function () {
	            return this._defaultCachingStore;
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(RuntimeConfigImpl.prototype, "defaultCachingTimeoutSeconds", {
	        get: function () {
	            return this._defaultCachingTimeoutSeconds;
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(RuntimeConfigImpl.prototype, "globalCacheDisable", {
	        get: function () {
	            return this._globalCacheDisable;
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(RuntimeConfigImpl.prototype, "fetchClientFactory", {
	        get: function () {
	            return this._fetchClientFactory;
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(RuntimeConfigImpl.prototype, "baseUrl", {
	        get: function () {
	            if (this._baseUrl !== null) {
	                return this._baseUrl;
	            }
	            else if (this._spfxContext !== null) {
	                return this._spfxContext.pageContext.web.absoluteUrl;
	            }
	            return null;
	        },
	        enumerable: true,
	        configurable: true
	    });
	    return RuntimeConfigImpl;
	}());
	exports.RuntimeConfigImpl = RuntimeConfigImpl;
	var _runtimeConfig = new RuntimeConfigImpl();
	exports.RuntimeConfig = _runtimeConfig;
	function setRuntimeConfig(config) {
	    _runtimeConfig.set(config);
	}
	exports.setRuntimeConfig = setRuntimeConfig;


/***/ },
/* 16 */
/***/ function(module, exports) {

	/* WEBPACK VAR INJECTION */(function(global) {"use strict";
	Object.defineProperty(exports, "__esModule", { value: true });
	/**
	 * Makes requests using the fetch API
	 */
	var FetchClient = (function () {
	    function FetchClient() {
	    }
	    FetchClient.prototype.fetch = function (url, options) {
	        return global.fetch(url, options);
	    };
	    return FetchClient;
	}());
	exports.FetchClient = FetchClient;
	
	/* WEBPACK VAR INJECTION */}.call(exports, (function() { return this; }())))

/***/ },
/* 17 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	Object.defineProperty(exports, "__esModule", { value: true });
	var util_1 = __webpack_require__(12);
	var collections_1 = __webpack_require__(18);
	var pnplibconfig_1 = __webpack_require__(15);
	/**
	 * A wrapper class to provide a consistent interface to browser based storage
	 *
	 */
	var PnPClientStorageWrapper = (function () {
	    /**
	     * Creates a new instance of the PnPClientStorageWrapper class
	     *
	     * @constructor
	     */
	    function PnPClientStorageWrapper(store, defaultTimeoutMinutes) {
	        this.store = store;
	        this.defaultTimeoutMinutes = defaultTimeoutMinutes;
	        this.defaultTimeoutMinutes = (defaultTimeoutMinutes === void 0) ? -1 : defaultTimeoutMinutes;
	        this.enabled = this.test();
	    }
	    /**
	     * Get a value from storage, or null if that value does not exist
	     *
	     * @param key The key whose value we want to retrieve
	     */
	    PnPClientStorageWrapper.prototype.get = function (key) {
	        if (!this.enabled) {
	            return null;
	        }
	        var o = this.store.getItem(key);
	        if (o == null) {
	            return null;
	        }
	        var persistable = JSON.parse(o);
	        if (new Date(persistable.expiration) <= new Date()) {
	            this.delete(key);
	            return null;
	        }
	        else {
	            return persistable.value;
	        }
	    };
	    /**
	     * Adds a value to the underlying storage
	     *
	     * @param key The key to use when storing the provided value
	     * @param o The value to store
	     * @param expire Optional, if provided the expiration of the item, otherwise the default is used
	     */
	    PnPClientStorageWrapper.prototype.put = function (key, o, expire) {
	        if (this.enabled) {
	            this.store.setItem(key, this.createPersistable(o, expire));
	        }
	    };
	    /**
	     * Deletes a value from the underlying storage
	     *
	     * @param key The key of the pair we want to remove from storage
	     */
	    PnPClientStorageWrapper.prototype.delete = function (key) {
	        if (this.enabled) {
	            this.store.removeItem(key);
	        }
	    };
	    /**
	     * Gets an item from the underlying storage, or adds it if it does not exist using the supplied getter function
	     *
	     * @param key The key to use when storing the provided value
	     * @param getter A function which will upon execution provide the desired value
	     * @param expire Optional, if provided the expiration of the item, otherwise the default is used
	     */
	    PnPClientStorageWrapper.prototype.getOrPut = function (key, getter, expire) {
	        var _this = this;
	        if (!this.enabled) {
	            return getter();
	        }
	        return new Promise(function (resolve) {
	            var o = _this.get(key);
	            if (o == null) {
	                getter().then(function (d) {
	                    _this.put(key, d, expire);
	                    resolve(d);
	                });
	            }
	            else {
	                resolve(o);
	            }
	        });
	    };
	    /**
	     * Used to determine if the wrapped storage is available currently
	     */
	    PnPClientStorageWrapper.prototype.test = function () {
	        var str = "test";
	        try {
	            this.store.setItem(str, str);
	            this.store.removeItem(str);
	            return true;
	        }
	        catch (e) {
	            return false;
	        }
	    };
	    /**
	     * Creates the persistable to store
	     */
	    PnPClientStorageWrapper.prototype.createPersistable = function (o, expire) {
	        if (typeof expire === "undefined") {
	            // ensure we are by default inline with the global library setting
	            var defaultTimeout = pnplibconfig_1.RuntimeConfig.defaultCachingTimeoutSeconds;
	            if (this.defaultTimeoutMinutes > 0) {
	                defaultTimeout = this.defaultTimeoutMinutes * 60;
	            }
	            expire = util_1.Util.dateAdd(new Date(), "second", defaultTimeout);
	        }
	        return JSON.stringify({ expiration: expire, value: o });
	    };
	    return PnPClientStorageWrapper;
	}());
	exports.PnPClientStorageWrapper = PnPClientStorageWrapper;
	/**
	 * A thin implementation of in-memory storage for use in nodejs
	 */
	var MemoryStorage = (function () {
	    function MemoryStorage(_store) {
	        if (_store === void 0) { _store = new collections_1.Dictionary(); }
	        this._store = _store;
	    }
	    Object.defineProperty(MemoryStorage.prototype, "length", {
	        get: function () {
	            return this._store.count();
	        },
	        enumerable: true,
	        configurable: true
	    });
	    MemoryStorage.prototype.clear = function () {
	        this._store.clear();
	    };
	    MemoryStorage.prototype.getItem = function (key) {
	        return this._store.get(key);
	    };
	    MemoryStorage.prototype.key = function (index) {
	        return this._store.getKeys()[index];
	    };
	    MemoryStorage.prototype.removeItem = function (key) {
	        this._store.remove(key);
	    };
	    MemoryStorage.prototype.setItem = function (key, data) {
	        this._store.add(key, data);
	    };
	    return MemoryStorage;
	}());
	/**
	 * A class that will establish wrappers for both local and session storage
	 */
	var PnPClientStorage = (function () {
	    /**
	     * Creates a new instance of the PnPClientStorage class
	     *
	     * @constructor
	     */
	    function PnPClientStorage() {
	        this.local = typeof localStorage !== "undefined" ? new PnPClientStorageWrapper(localStorage) : new PnPClientStorageWrapper(new MemoryStorage());
	        this.session = typeof sessionStorage !== "undefined" ? new PnPClientStorageWrapper(sessionStorage) : new PnPClientStorageWrapper(new MemoryStorage());
	    }
	    return PnPClientStorage;
	}());
	exports.PnPClientStorage = PnPClientStorage;


/***/ },
/* 18 */
/***/ function(module, exports) {

	"use strict";
	Object.defineProperty(exports, "__esModule", { value: true });
	/**
	 * Generic dictionary
	 */
	var Dictionary = (function () {
	    /**
	     * Creates a new instance of the Dictionary<T> class
	     *
	     * @constructor
	     */
	    function Dictionary(keys, values) {
	        if (keys === void 0) { keys = []; }
	        if (values === void 0) { values = []; }
	        this.keys = keys;
	        this.values = values;
	    }
	    /**
	     * Gets a value from the collection using the specified key
	     *
	     * @param key The key whose value we want to return, returns null if the key does not exist
	     */
	    Dictionary.prototype.get = function (key) {
	        var index = this.keys.indexOf(key);
	        if (index < 0) {
	            return null;
	        }
	        return this.values[index];
	    };
	    /**
	     * Adds the supplied key and value to the dictionary
	     *
	     * @param key The key to add
	     * @param o The value to add
	     */
	    Dictionary.prototype.add = function (key, o) {
	        var index = this.keys.indexOf(key);
	        if (index > -1) {
	            this.values[index] = o;
	        }
	        else {
	            this.keys.push(key);
	            this.values.push(o);
	        }
	    };
	    /**
	     * Merges the supplied typed hash into this dictionary instance. Existing values are updated and new ones are created as appropriate.
	     */
	    Dictionary.prototype.merge = function (source) {
	        var _this = this;
	        if ("getKeys" in source) {
	            var sourceAsDictionary_1 = source;
	            sourceAsDictionary_1.getKeys().map(function (key) {
	                _this.add(key, sourceAsDictionary_1.get(key));
	            });
	        }
	        else {
	            var sourceAsHash = source;
	            for (var key in sourceAsHash) {
	                if (sourceAsHash.hasOwnProperty(key)) {
	                    this.add(key, sourceAsHash[key]);
	                }
	            }
	        }
	    };
	    /**
	     * Removes a value from the dictionary
	     *
	     * @param key The key of the key/value pair to remove. Returns null if the key was not found.
	     */
	    Dictionary.prototype.remove = function (key) {
	        var index = this.keys.indexOf(key);
	        if (index < 0) {
	            return null;
	        }
	        var val = this.values[index];
	        this.keys.splice(index, 1);
	        this.values.splice(index, 1);
	        return val;
	    };
	    /**
	     * Returns all the keys currently in the dictionary as an array
	     */
	    Dictionary.prototype.getKeys = function () {
	        return this.keys;
	    };
	    /**
	     * Returns all the values currently in the dictionary as an array
	     */
	    Dictionary.prototype.getValues = function () {
	        return this.values;
	    };
	    /**
	     * Clears the current dictionary
	     */
	    Dictionary.prototype.clear = function () {
	        this.keys = [];
	        this.values = [];
	    };
	    /**
	     * Gets a count of the items currently in the dictionary
	     */
	    Dictionary.prototype.count = function () {
	        return this.keys.length;
	    };
	    return Dictionary;
	}());
	exports.Dictionary = Dictionary;


/***/ },
/* 19 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	Object.defineProperty(exports, "__esModule", { value: true });
	var collections_1 = __webpack_require__(18);
	/**
	 * Class used to manage the current application settings
	 *
	 */
	var Settings = (function () {
	    /**
	     * Creates a new instance of the settings class
	     *
	     * @constructor
	     */
	    function Settings() {
	        this._settings = new collections_1.Dictionary();
	    }
	    /**
	     * Adds a new single setting, or overwrites a previous setting with the same key
	     *
	     * @param {string} key The key used to store this setting
	     * @param {string} value The setting value to store
	     */
	    Settings.prototype.add = function (key, value) {
	        this._settings.add(key, value);
	    };
	    /**
	     * Adds a JSON value to the collection as a string, you must use getJSON to rehydrate the object when read
	     *
	     * @param {string} key The key used to store this setting
	     * @param {any} value The setting value to store
	     */
	    Settings.prototype.addJSON = function (key, value) {
	        this._settings.add(key, JSON.stringify(value));
	    };
	    /**
	     * Applies the supplied hash to the setting collection overwriting any existing value, or created new values
	     *
	     * @param {TypedHash<any>} hash The set of values to add
	     */
	    Settings.prototype.apply = function (hash) {
	        var _this = this;
	        return new Promise(function (resolve, reject) {
	            try {
	                _this._settings.merge(hash);
	                resolve();
	            }
	            catch (e) {
	                reject(e);
	            }
	        });
	    };
	    /**
	     * Loads configuration settings into the collection from the supplied provider and returns a Promise
	     *
	     * @param {IConfigurationProvider} provider The provider from which we will load the settings
	     */
	    Settings.prototype.load = function (provider) {
	        var _this = this;
	        return new Promise(function (resolve, reject) {
	            provider.getConfiguration().then(function (value) {
	                _this._settings.merge(value);
	                resolve();
	            }).catch(function (reason) {
	                reject(reason);
	            });
	        });
	    };
	    /**
	     * Gets a value from the configuration
	     *
	     * @param {string} key The key whose value we want to return. Returns null if the key does not exist
	     * @return {string} string value from the configuration
	     */
	    Settings.prototype.get = function (key) {
	        return this._settings.get(key);
	    };
	    /**
	     * Gets a JSON value, rehydrating the stored string to the original object
	     *
	     * @param {string} key The key whose value we want to return. Returns null if the key does not exist
	     * @return {any} object from the configuration
	     */
	    Settings.prototype.getJSON = function (key) {
	        var o = this.get(key);
	        if (typeof o === "undefined" || o === null) {
	            return o;
	        }
	        return JSON.parse(o);
	    };
	    return Settings;
	}());
	exports.Settings = Settings;


/***/ },
/* 20 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	Object.defineProperty(exports, "__esModule", { value: true });
	var search_1 = __webpack_require__(21);
	var searchsuggest_1 = __webpack_require__(29);
	var site_1 = __webpack_require__(30);
	var webs_1 = __webpack_require__(31);
	var util_1 = __webpack_require__(12);
	var userprofiles_1 = __webpack_require__(53);
	var exceptions_1 = __webpack_require__(26);
	/**
	 * Root of the SharePoint REST module
	 */
	var Rest = (function () {
	    function Rest() {
	    }
	    /**
	     * Executes a search against this web context
	     *
	     * @param query The SearchQuery definition
	     */
	    Rest.prototype.searchSuggest = function (query) {
	        var finalQuery;
	        if (typeof query === "string") {
	            finalQuery = { querytext: query };
	        }
	        else {
	            finalQuery = query;
	        }
	        return new searchsuggest_1.SearchSuggest("").execute(finalQuery);
	    };
	    /**
	     * Executes a search against this web context
	     *
	     * @param query The SearchQuery definition
	     */
	    Rest.prototype.search = function (query) {
	        var finalQuery;
	        if (typeof query === "string") {
	            finalQuery = { Querytext: query };
	        }
	        else {
	            finalQuery = query;
	        }
	        return new search_1.Search("").execute(finalQuery);
	    };
	    Object.defineProperty(Rest.prototype, "site", {
	        /**
	         * Begins a site collection scoped REST request
	         *
	         */
	        get: function () {
	            return new site_1.Site("");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Rest.prototype, "web", {
	        /**
	         * Begins a web scoped REST request
	         *
	         */
	        get: function () {
	            return new webs_1.Web("");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Rest.prototype, "profiles", {
	        /**
	         * Access to user profile methods
	         *
	         */
	        get: function () {
	            return new userprofiles_1.UserProfileQuery("");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    /**
	     * Creates a new batch object for use with the Queryable.addToBatch method
	     *
	     */
	    Rest.prototype.createBatch = function () {
	        return this.web.createBatch();
	    };
	    /**
	     * Begins a cross-domain, host site scoped REST request, for use in add-in webs
	     *
	     * @param addInWebUrl The absolute url of the add-in web
	     * @param hostWebUrl The absolute url of the host web
	     */
	    Rest.prototype.crossDomainSite = function (addInWebUrl, hostWebUrl) {
	        return this._cdImpl(site_1.Site, addInWebUrl, hostWebUrl, "site");
	    };
	    /**
	     * Begins a cross-domain, host web scoped REST request, for use in add-in webs
	     *
	     * @param addInWebUrl The absolute url of the add-in web
	     * @param hostWebUrl The absolute url of the host web
	     */
	    Rest.prototype.crossDomainWeb = function (addInWebUrl, hostWebUrl) {
	        return this._cdImpl(webs_1.Web, addInWebUrl, hostWebUrl, "web");
	    };
	    /**
	     * Implements the creation of cross domain REST urls
	     *
	     * @param factory The constructor of the object to create Site | Web
	     * @param addInWebUrl The absolute url of the add-in web
	     * @param hostWebUrl The absolute url of the host web
	     * @param urlPart String part to append to the url "site" | "web"
	     */
	    Rest.prototype._cdImpl = function (factory, addInWebUrl, hostWebUrl, urlPart) {
	        if (!util_1.Util.isUrlAbsolute(addInWebUrl)) {
	            throw new exceptions_1.UrlException("The addInWebUrl parameter must be an absolute url.");
	        }
	        if (!util_1.Util.isUrlAbsolute(hostWebUrl)) {
	            throw new exceptions_1.UrlException("The hostWebUrl parameter must be an absolute url.");
	        }
	        var url = util_1.Util.combinePaths(addInWebUrl, "_api/SP.AppContextSite(@target)");
	        var instance = new factory(url, urlPart);
	        instance.query.add("@target", "'" + encodeURIComponent(hostWebUrl) + "'");
	        return instance;
	    };
	    return Rest;
	}());
	exports.Rest = Rest;


/***/ },
/* 21 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	var __extends = (this && this.__extends) || (function () {
	    var extendStatics = Object.setPrototypeOf ||
	        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
	        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
	    return function (d, b) {
	        extendStatics(d, b);
	        function __() { this.constructor = d; }
	        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
	    };
	})();
	Object.defineProperty(exports, "__esModule", { value: true });
	var queryable_1 = __webpack_require__(22);
	var util_1 = __webpack_require__(12);
	/**
	 * Describes the search API
	 *
	 */
	var Search = (function (_super) {
	    __extends(Search, _super);
	    /**
	     * Creates a new instance of the Search class
	     *
	     * @param baseUrl The url for the search context
	     * @param query The SearchQuery object to execute
	     */
	    function Search(baseUrl, path) {
	        if (path === void 0) { path = "_api/search/postquery"; }
	        return _super.call(this, baseUrl, path) || this;
	    }
	    /**
	     * .......
	     * @returns Promise
	     */
	    Search.prototype.execute = function (query) {
	        var formattedBody;
	        formattedBody = query;
	        if (formattedBody.SelectProperties) {
	            formattedBody.SelectProperties = { results: query.SelectProperties };
	        }
	        if (formattedBody.RefinementFilters) {
	            formattedBody.RefinementFilters = { results: query.RefinementFilters };
	        }
	        if (formattedBody.SortList) {
	            formattedBody.SortList = { results: query.SortList };
	        }
	        if (formattedBody.HithighlightedProperties) {
	            formattedBody.HithighlightedProperties = { results: query.HithighlightedProperties };
	        }
	        if (formattedBody.ReorderingRules) {
	            formattedBody.ReorderingRules = { results: query.ReorderingRules };
	        }
	        if (formattedBody.Properties) {
	            formattedBody.Properties = { results: query.Properties };
	        }
	        var postBody = JSON.stringify({
	            request: util_1.Util.extend({
	                "__metadata": { "type": "Microsoft.Office.Server.Search.REST.SearchRequest" },
	            }, formattedBody),
	        });
	        return this.post({ body: postBody }).then(function (data) { return new SearchResults(data); });
	    };
	    return Search;
	}(queryable_1.QueryableInstance));
	exports.Search = Search;
	/**
	 * Describes the SearchResults class, which returns the formatted and raw version of the query response
	 */
	var SearchResults = (function () {
	    /**
	     * Creates a new instance of the SearchResult class
	     *
	     */
	    function SearchResults(rawResponse) {
	        var response = rawResponse.postquery ? rawResponse.postquery : rawResponse;
	        this.PrimarySearchResults = this.formatSearchResults(response.PrimaryQueryResult.RelevantResults.Table.Rows);
	        this.RawSearchResults = response;
	        this.ElapsedTime = response.ElapsedTime;
	        this.RowCount = response.PrimaryQueryResult.RelevantResults.RowCount;
	        this.TotalRows = response.PrimaryQueryResult.RelevantResults.TotalRows;
	        this.TotalRowsIncludingDuplicates = response.PrimaryQueryResult.RelevantResults.TotalRowsIncludingDuplicates;
	    }
	    /**
	     * Formats a search results array
	     *
	     * @param rawResults The array to process
	     */
	    SearchResults.prototype.formatSearchResults = function (rawResults) {
	        var results = new Array(), tempResults = rawResults.results ? rawResults.results : rawResults;
	        for (var _i = 0, tempResults_1 = tempResults; _i < tempResults_1.length; _i++) {
	            var i = tempResults_1[_i];
	            results.push(new SearchResult(i.Cells));
	        }
	        return results;
	    };
	    return SearchResults;
	}());
	exports.SearchResults = SearchResults;
	/**
	 * Describes the SearchResult class
	 */
	var SearchResult = (function () {
	    /**
	     * Creates a new instance of the SearchResult class
	     *
	     */
	    function SearchResult(rawItem) {
	        var item = rawItem.results ? rawItem.results : rawItem;
	        for (var _i = 0, item_1 = item; _i < item_1.length; _i++) {
	            var i = item_1[_i];
	            Object.defineProperty(this, i.Key, {
	                configurable: false,
	                enumerable: false,
	                value: i.Value,
	                writable: false,
	            });
	        }
	    }
	    return SearchResult;
	}());
	exports.SearchResult = SearchResult;
	/**
	 * defines the SortDirection enum
	 */
	var SortDirection;
	(function (SortDirection) {
	    SortDirection[SortDirection["Ascending"] = 0] = "Ascending";
	    SortDirection[SortDirection["Descending"] = 1] = "Descending";
	    SortDirection[SortDirection["FQLFormula"] = 2] = "FQLFormula";
	})(SortDirection = exports.SortDirection || (exports.SortDirection = {}));
	/**
	 * defines the ReorderingRuleMatchType  enum
	 */
	var ReorderingRuleMatchType;
	(function (ReorderingRuleMatchType) {
	    ReorderingRuleMatchType[ReorderingRuleMatchType["ResultContainsKeyword"] = 0] = "ResultContainsKeyword";
	    ReorderingRuleMatchType[ReorderingRuleMatchType["TitleContainsKeyword"] = 1] = "TitleContainsKeyword";
	    ReorderingRuleMatchType[ReorderingRuleMatchType["TitleMatchesKeyword"] = 2] = "TitleMatchesKeyword";
	    ReorderingRuleMatchType[ReorderingRuleMatchType["UrlStartsWith"] = 3] = "UrlStartsWith";
	    ReorderingRuleMatchType[ReorderingRuleMatchType["UrlExactlyMatches"] = 4] = "UrlExactlyMatches";
	    ReorderingRuleMatchType[ReorderingRuleMatchType["ContentTypeIs"] = 5] = "ContentTypeIs";
	    ReorderingRuleMatchType[ReorderingRuleMatchType["FileExtensionMatches"] = 6] = "FileExtensionMatches";
	    ReorderingRuleMatchType[ReorderingRuleMatchType["ResultHasTag"] = 7] = "ResultHasTag";
	    ReorderingRuleMatchType[ReorderingRuleMatchType["ManualCondition"] = 8] = "ManualCondition";
	})(ReorderingRuleMatchType = exports.ReorderingRuleMatchType || (exports.ReorderingRuleMatchType = {}));
	/**
	 * Specifies the type value for the property
	 */
	var QueryPropertyValueType;
	(function (QueryPropertyValueType) {
	    QueryPropertyValueType[QueryPropertyValueType["None"] = 0] = "None";
	    QueryPropertyValueType[QueryPropertyValueType["StringType"] = 1] = "StringType";
	    QueryPropertyValueType[QueryPropertyValueType["Int32TYpe"] = 2] = "Int32TYpe";
	    QueryPropertyValueType[QueryPropertyValueType["BooleanType"] = 3] = "BooleanType";
	    QueryPropertyValueType[QueryPropertyValueType["StringArrayType"] = 4] = "StringArrayType";
	    QueryPropertyValueType[QueryPropertyValueType["UnSupportedType"] = 5] = "UnSupportedType";
	})(QueryPropertyValueType = exports.QueryPropertyValueType || (exports.QueryPropertyValueType = {}));


/***/ },
/* 22 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	var __extends = (this && this.__extends) || (function () {
	    var extendStatics = Object.setPrototypeOf ||
	        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
	        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
	    return function (d, b) {
	        extendStatics(d, b);
	        function __() { this.constructor = d; }
	        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
	    };
	})();
	Object.defineProperty(exports, "__esModule", { value: true });
	var util_1 = __webpack_require__(12);
	var collections_1 = __webpack_require__(18);
	var odata_1 = __webpack_require__(23);
	var pnplibconfig_1 = __webpack_require__(15);
	var exceptions_1 = __webpack_require__(26);
	var logging_1 = __webpack_require__(14);
	var pipeline_1 = __webpack_require__(27);
	/**
	 * Queryable Base Class
	 *
	 */
	var Queryable = (function () {
	    /**
	     * Creates a new instance of the Queryable class
	     *
	     * @constructor
	     * @param baseUrl A string or Queryable that should form the base part of the url
	     *
	     */
	    function Queryable(baseUrl, path) {
	        this._query = new collections_1.Dictionary();
	        this._batch = null;
	        if (typeof baseUrl === "string") {
	            // we need to do some extra parsing to get the parent url correct if we are
	            // being created from just a string.
	            var urlStr = baseUrl;
	            if (util_1.Util.isUrlAbsolute(urlStr) || urlStr.lastIndexOf("/") < 0) {
	                this._parentUrl = urlStr;
	                this._url = util_1.Util.combinePaths(urlStr, path);
	            }
	            else if (urlStr.lastIndexOf("/") > urlStr.lastIndexOf("(")) {
	                // .../items(19)/fields
	                var index = urlStr.lastIndexOf("/");
	                this._parentUrl = urlStr.slice(0, index);
	                path = util_1.Util.combinePaths(urlStr.slice(index), path);
	                this._url = util_1.Util.combinePaths(this._parentUrl, path);
	            }
	            else {
	                // .../items(19)
	                var index = urlStr.lastIndexOf("(");
	                this._parentUrl = urlStr.slice(0, index);
	                this._url = util_1.Util.combinePaths(urlStr, path);
	            }
	        }
	        else {
	            var q = baseUrl;
	            this._parentUrl = q._url;
	            var target = q._query.get("@target");
	            if (target !== null) {
	                this._query.add("@target", target);
	            }
	            this._url = util_1.Util.combinePaths(this._parentUrl, path);
	        }
	    }
	    /**
	     * Directly concatonates the supplied string to the current url, not normalizing "/" chars
	     *
	     * @param pathPart The string to concatonate to the url
	     */
	    Queryable.prototype.concat = function (pathPart) {
	        this._url += pathPart;
	        return this;
	    };
	    /**
	     * Appends the given string and normalizes "/" chars
	     *
	     * @param pathPart The string to append
	     */
	    Queryable.prototype.append = function (pathPart) {
	        this._url = util_1.Util.combinePaths(this._url, pathPart);
	    };
	    /**
	     * Blocks a batch call from occuring, MUST be cleared by calling the returned function
	     */
	    Queryable.prototype.addBatchDependency = function () {
	        if (this.hasBatch) {
	            return this._batch.addDependency();
	        }
	        return function () { return null; };
	    };
	    Object.defineProperty(Queryable.prototype, "hasBatch", {
	        /**
	         * Indicates if the current query has a batch associated
	         *
	         */
	        get: function () {
	            return this._batch !== null;
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Queryable.prototype, "batch", {
	        /**
	         * The batch currently associated with this query or null
	         *
	         */
	        get: function () {
	            return this.hasBatch ? this._batch : null;
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Queryable.prototype, "parentUrl", {
	        /**
	         * Gets the parent url used when creating this instance
	         *
	         */
	        get: function () {
	            return this._parentUrl;
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Queryable.prototype, "query", {
	        /**
	         * Provides access to the query builder for this url
	         *
	         */
	        get: function () {
	            return this._query;
	        },
	        enumerable: true,
	        configurable: true
	    });
	    /**
	     * Creates a new instance of the supplied factory and extends this into that new instance
	     *
	     * @param factory constructor for the new queryable
	     */
	    Queryable.prototype.as = function (factory) {
	        var o = new factory(this._url, null);
	        return util_1.Util.extend(o, this, true);
	    };
	    /**
	     * Adds this query to the supplied batch
	     *
	     * @example
	     * ```
	     *
	     * let b = pnp.sp.createBatch();
	     * pnp.sp.web.inBatch(b).get().then(...);
	     * b.execute().then(...)
	     * ```
	     */
	    Queryable.prototype.inBatch = function (batch) {
	        if (this._batch !== null) {
	            throw new exceptions_1.AlreadyInBatchException();
	        }
	        this._batch = batch;
	        return this;
	    };
	    /**
	     * Enables caching for this request
	     *
	     * @param options Defines the options used when caching this request
	     */
	    Queryable.prototype.usingCaching = function (options) {
	        if (!pnplibconfig_1.RuntimeConfig.globalCacheDisable) {
	            this._useCaching = true;
	            this._cachingOptions = options;
	        }
	        return this;
	    };
	    /**
	     * Gets the currentl url, made absolute based on the availability of the _spPageContextInfo object
	     *
	     */
	    Queryable.prototype.toUrl = function () {
	        return this._url;
	    };
	    /**
	     * Gets the full url with query information
	     *
	     */
	    Queryable.prototype.toUrlAndQuery = function () {
	        var aliasedParams = new collections_1.Dictionary();
	        var url = this.toUrl().replace(/'!(@.*?)::(.*?)'/ig, function (match, labelName, value) {
	            logging_1.Logger.write("Rewriting aliased parameter from match " + match + " to label: " + labelName + " value: " + value, logging_1.LogLevel.Verbose);
	            aliasedParams.add(labelName, "'" + value + "'");
	            return labelName;
	        });
	        // inlude our explicitly set query string params
	        aliasedParams.merge(this._query);
	        if (aliasedParams.count() > 0) {
	            url += "?" + aliasedParams.getKeys().map(function (key) { return key + "=" + aliasedParams.get(key); }).join("&");
	        }
	        return url;
	    };
	    /**
	     * Gets a parent for this instance as specified
	     *
	     * @param factory The contructor for the class to create
	     */
	    Queryable.prototype.getParent = function (factory, baseUrl, path, batch) {
	        if (baseUrl === void 0) { baseUrl = this.parentUrl; }
	        var parent = new factory(baseUrl, path);
	        var target = this.query.get("@target");
	        if (target !== null) {
	            parent.query.add("@target", target);
	        }
	        if (typeof batch !== "undefined") {
	            parent = parent.inBatch(batch);
	        }
	        return parent;
	    };
	    /**
	     * Clones this queryable into a new queryable instance of T
	     * @param factory Constructor used to create the new instance
	     * @param additionalPath Any additional path to include in the clone
	     * @param includeBatch If true this instance's batch will be added to the cloned instance
	     */
	    Queryable.prototype.clone = function (factory, additionalPath, includeBatch) {
	        if (includeBatch === void 0) { includeBatch = false; }
	        var clone = new factory(this, additionalPath);
	        var target = this.query.get("@target");
	        if (target !== null) {
	            clone.query.add("@target", target);
	        }
	        if (includeBatch && this.hasBatch) {
	            clone = clone.inBatch(this.batch);
	        }
	        return clone;
	    };
	    /**
	     * Executes the currently built request
	     *
	     * @param parser Allows you to specify a parser to handle the result
	     * @param getOptions The options used for this request
	     */
	    Queryable.prototype.get = function (parser, getOptions) {
	        if (parser === void 0) { parser = new odata_1.ODataDefaultParser(); }
	        if (getOptions === void 0) { getOptions = {}; }
	        return this.toRequestContext("GET", getOptions, parser).then(function (context) { return pipeline_1.pipe(context); });
	    };
	    Queryable.prototype.getAs = function (parser, getOptions) {
	        if (parser === void 0) { parser = new odata_1.ODataDefaultParser(); }
	        if (getOptions === void 0) { getOptions = {}; }
	        return this.toRequestContext("GET", getOptions, parser).then(function (context) { return pipeline_1.pipe(context); });
	    };
	    Queryable.prototype.post = function (postOptions, parser) {
	        if (postOptions === void 0) { postOptions = {}; }
	        if (parser === void 0) { parser = new odata_1.ODataDefaultParser(); }
	        return this.toRequestContext("POST", postOptions, parser).then(function (context) { return pipeline_1.pipe(context); });
	    };
	    Queryable.prototype.postAs = function (postOptions, parser) {
	        if (postOptions === void 0) { postOptions = {}; }
	        if (parser === void 0) { parser = new odata_1.ODataDefaultParser(); }
	        return this.toRequestContext("POST", postOptions, parser).then(function (context) { return pipeline_1.pipe(context); });
	    };
	    Queryable.prototype.patch = function (patchOptions, parser) {
	        if (patchOptions === void 0) { patchOptions = {}; }
	        if (parser === void 0) { parser = new odata_1.ODataDefaultParser(); }
	        return this.toRequestContext("PATCH", patchOptions, parser).then(function (context) { return pipeline_1.pipe(context); });
	    };
	    Queryable.prototype.delete = function (deleteOptions, parser) {
	        if (deleteOptions === void 0) { deleteOptions = {}; }
	        if (parser === void 0) { parser = new odata_1.ODataDefaultParser(); }
	        return this.toRequestContext("DELETE", deleteOptions, parser).then(function (context) { return pipeline_1.pipe(context); });
	    };
	    /**
	     * Converts the current instance to a request context
	     *
	     * @param verb The request verb
	     * @param options The set of supplied request options
	     * @param parser The supplied ODataParser instance
	     * @param pipeline Optional request processing pipeline
	     */
	    Queryable.prototype.toRequestContext = function (verb, options, parser, pipeline) {
	        var _this = this;
	        if (options === void 0) { options = {}; }
	        if (pipeline === void 0) { pipeline = pipeline_1.PipelineMethods.default; }
	        var dependencyDispose = this.hasBatch ? this.addBatchDependency() : function () { return; };
	        return util_1.Util.toAbsoluteUrl(this.toUrlAndQuery()).then(function (url) {
	            // build our request context
	            var context = {
	                batch: _this._batch,
	                batchDependency: dependencyDispose,
	                cachingOptions: _this._cachingOptions,
	                isBatched: _this.hasBatch,
	                isCached: _this._useCaching,
	                options: options,
	                parser: parser,
	                pipeline: pipeline,
	                requestAbsoluteUrl: url,
	                requestId: util_1.Util.getGUID(),
	                verb: verb,
	            };
	            return context;
	        });
	    };
	    return Queryable;
	}());
	exports.Queryable = Queryable;
	/**
	 * Represents a REST collection which can be filtered, paged, and selected
	 *
	 */
	var QueryableCollection = (function (_super) {
	    __extends(QueryableCollection, _super);
	    function QueryableCollection() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    /**
	     * Filters the returned collection (https://msdn.microsoft.com/en-us/library/office/fp142385.aspx#bk_supported)
	     *
	     * @param filter The string representing the filter query
	     */
	    QueryableCollection.prototype.filter = function (filter) {
	        this._query.add("$filter", filter);
	        return this;
	    };
	    /**
	     * Choose which fields to return
	     *
	     * @param selects One or more fields to return
	     */
	    QueryableCollection.prototype.select = function () {
	        var selects = [];
	        for (var _i = 0; _i < arguments.length; _i++) {
	            selects[_i] = arguments[_i];
	        }
	        if (selects.length > 0) {
	            this._query.add("$select", selects.join(","));
	        }
	        return this;
	    };
	    /**
	     * Expands fields such as lookups to get additional data
	     *
	     * @param expands The Fields for which to expand the values
	     */
	    QueryableCollection.prototype.expand = function () {
	        var expands = [];
	        for (var _i = 0; _i < arguments.length; _i++) {
	            expands[_i] = arguments[_i];
	        }
	        if (expands.length > 0) {
	            this._query.add("$expand", expands.join(","));
	        }
	        return this;
	    };
	    /**
	     * Orders based on the supplied fields ascending
	     *
	     * @param orderby The name of the field to sort on
	     * @param ascending If false DESC is appended, otherwise ASC (default)
	     */
	    QueryableCollection.prototype.orderBy = function (orderBy, ascending) {
	        if (ascending === void 0) { ascending = true; }
	        var keys = this._query.getKeys();
	        var query = [];
	        var asc = ascending ? " asc" : " desc";
	        for (var i = 0; i < keys.length; i++) {
	            if (keys[i] === "$orderby") {
	                query.push(this._query.get("$orderby"));
	                break;
	            }
	        }
	        query.push("" + orderBy + asc);
	        this._query.add("$orderby", query.join(","));
	        return this;
	    };
	    /**
	     * Skips the specified number of items
	     *
	     * @param skip The number of items to skip
	     */
	    QueryableCollection.prototype.skip = function (skip) {
	        this._query.add("$skip", skip.toString());
	        return this;
	    };
	    /**
	     * Limits the query to only return the specified number of items
	     *
	     * @param top The query row limit
	     */
	    QueryableCollection.prototype.top = function (top) {
	        this._query.add("$top", top.toString());
	        return this;
	    };
	    return QueryableCollection;
	}(Queryable));
	exports.QueryableCollection = QueryableCollection;
	/**
	 * Represents an instance that can be selected
	 *
	 */
	var QueryableInstance = (function (_super) {
	    __extends(QueryableInstance, _super);
	    function QueryableInstance() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    /**
	     * Choose which fields to return
	     *
	     * @param selects One or more fields to return
	     */
	    QueryableInstance.prototype.select = function () {
	        var selects = [];
	        for (var _i = 0; _i < arguments.length; _i++) {
	            selects[_i] = arguments[_i];
	        }
	        if (selects.length > 0) {
	            this._query.add("$select", selects.join(","));
	        }
	        return this;
	    };
	    /**
	     * Expands fields such as lookups to get additional data
	     *
	     * @param expands The Fields for which to expand the values
	     */
	    QueryableInstance.prototype.expand = function () {
	        var expands = [];
	        for (var _i = 0; _i < arguments.length; _i++) {
	            expands[_i] = arguments[_i];
	        }
	        if (expands.length > 0) {
	            this._query.add("$expand", expands.join(","));
	        }
	        return this;
	    };
	    return QueryableInstance;
	}(Queryable));
	exports.QueryableInstance = QueryableInstance;


/***/ },
/* 23 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	var __extends = (this && this.__extends) || (function () {
	    var extendStatics = Object.setPrototypeOf ||
	        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
	        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
	    return function (d, b) {
	        extendStatics(d, b);
	        function __() { this.constructor = d; }
	        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
	    };
	})();
	Object.defineProperty(exports, "__esModule", { value: true });
	var util_1 = __webpack_require__(12);
	var logging_1 = __webpack_require__(14);
	var httpclient_1 = __webpack_require__(24);
	var pnplibconfig_1 = __webpack_require__(15);
	var exceptions_1 = __webpack_require__(26);
	var exceptions_2 = __webpack_require__(26);
	function extractOdataId(candidate) {
	    if (candidate.hasOwnProperty("odata.id")) {
	        return candidate["odata.id"];
	    }
	    else if (candidate.hasOwnProperty("__metadata") && candidate.__metadata.hasOwnProperty("id")) {
	        return candidate.__metadata.id;
	    }
	    else {
	        throw new exceptions_1.ODataIdException(candidate);
	    }
	}
	exports.extractOdataId = extractOdataId;
	var ODataParserBase = (function () {
	    function ODataParserBase() {
	    }
	    ODataParserBase.prototype.parse = function (r) {
	        var _this = this;
	        return new Promise(function (resolve, reject) {
	            if (_this.handleError(r, reject)) {
	                if ((r.headers.has("Content-Length") && parseFloat(r.headers.get("Content-Length")) === 0) || r.status === 204) {
	                    resolve({});
	                }
	                else {
	                    r.json().then(function (json) { return resolve(_this.parseODataJSON(json)); }).catch(function (e) { return reject(e); });
	                }
	            }
	        });
	    };
	    ODataParserBase.prototype.handleError = function (r, reject) {
	        if (!r.ok) {
	            r.json().then(function (json) {
	                // include the headers as they contain diagnostic information
	                var data = {
	                    responseBody: json,
	                    responseHeaders: r.headers,
	                };
	                reject(new exceptions_2.ProcessHttpClientResponseException(r.status, r.statusText, data));
	            }).catch(function (e) {
	                // we failed to read the body - possibly it is empty. Let's report the original status that caused
	                // the request to fail and log the error with parsing the body if anyone needs it for debugging
	                logging_1.Logger.log({
	                    data: e,
	                    level: logging_1.LogLevel.Warning,
	                    message: "There was an error parsing the error response body. See data for details.",
	                });
	                // include the headers as they contain diagnostic information
	                var data = {
	                    responseBody: "[[body not available]]",
	                    responseHeaders: r.headers,
	                };
	                reject(new exceptions_2.ProcessHttpClientResponseException(r.status, r.statusText, data));
	            });
	        }
	        return r.ok;
	    };
	    ODataParserBase.prototype.parseODataJSON = function (json) {
	        var result = json;
	        if (json.hasOwnProperty("d")) {
	            if (json.d.hasOwnProperty("results")) {
	                result = json.d.results;
	            }
	            else {
	                result = json.d;
	            }
	        }
	        else if (json.hasOwnProperty("value")) {
	            result = json.value;
	        }
	        return result;
	    };
	    return ODataParserBase;
	}());
	exports.ODataParserBase = ODataParserBase;
	var ODataDefaultParser = (function (_super) {
	    __extends(ODataDefaultParser, _super);
	    function ODataDefaultParser() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    return ODataDefaultParser;
	}(ODataParserBase));
	exports.ODataDefaultParser = ODataDefaultParser;
	var ODataRawParserImpl = (function () {
	    function ODataRawParserImpl() {
	    }
	    ODataRawParserImpl.prototype.parse = function (r) {
	        return r.json();
	    };
	    return ODataRawParserImpl;
	}());
	exports.ODataRawParserImpl = ODataRawParserImpl;
	var ODataValueParserImpl = (function (_super) {
	    __extends(ODataValueParserImpl, _super);
	    function ODataValueParserImpl() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    ODataValueParserImpl.prototype.parse = function (r) {
	        return _super.prototype.parse.call(this, r).then(function (d) { return d; });
	    };
	    return ODataValueParserImpl;
	}(ODataParserBase));
	var ODataEntityParserImpl = (function (_super) {
	    __extends(ODataEntityParserImpl, _super);
	    function ODataEntityParserImpl(factory) {
	        var _this = _super.call(this) || this;
	        _this.factory = factory;
	        return _this;
	    }
	    ODataEntityParserImpl.prototype.parse = function (r) {
	        var _this = this;
	        return _super.prototype.parse.call(this, r).then(function (d) {
	            var o = new _this.factory(getEntityUrl(d), null);
	            return util_1.Util.extend(o, d);
	        });
	    };
	    return ODataEntityParserImpl;
	}(ODataParserBase));
	var ODataEntityArrayParserImpl = (function (_super) {
	    __extends(ODataEntityArrayParserImpl, _super);
	    function ODataEntityArrayParserImpl(factory) {
	        var _this = _super.call(this) || this;
	        _this.factory = factory;
	        return _this;
	    }
	    ODataEntityArrayParserImpl.prototype.parse = function (r) {
	        var _this = this;
	        return _super.prototype.parse.call(this, r).then(function (d) {
	            return d.map(function (v) {
	                var o = new _this.factory(getEntityUrl(v), null);
	                return util_1.Util.extend(o, v);
	            });
	        });
	    };
	    return ODataEntityArrayParserImpl;
	}(ODataParserBase));
	function getEntityUrl(entity) {
	    if (entity.hasOwnProperty("odata.editLink")) {
	        // we are dealign with minimal metadata (default)
	        return util_1.Util.combinePaths("_api", entity["odata.editLink"]);
	    }
	    else if (entity.hasOwnProperty("__metadata")) {
	        // we are dealing with verbose, which has an absolute uri
	        return entity.__metadata.uri;
	    }
	    else {
	        // we are likely dealing with nometadata, so don't error but we won't be able to
	        // chain off these objects
	        logging_1.Logger.write("No uri information found in ODataEntity parsing, chaining will fail for this object.", logging_1.LogLevel.Warning);
	        return "";
	    }
	}
	exports.getEntityUrl = getEntityUrl;
	exports.ODataRaw = new ODataRawParserImpl();
	function ODataValue() {
	    return new ODataValueParserImpl();
	}
	exports.ODataValue = ODataValue;
	function ODataEntity(factory) {
	    return new ODataEntityParserImpl(factory);
	}
	exports.ODataEntity = ODataEntity;
	function ODataEntityArray(factory) {
	    return new ODataEntityArrayParserImpl(factory);
	}
	exports.ODataEntityArray = ODataEntityArray;
	/**
	 * Manages a batch of OData operations
	 */
	var ODataBatch = (function () {
	    function ODataBatch(baseUrl, _batchId) {
	        if (_batchId === void 0) { _batchId = util_1.Util.getGUID(); }
	        this.baseUrl = baseUrl;
	        this._batchId = _batchId;
	        this._requests = [];
	        this._dependencies = [];
	    }
	    Object.defineProperty(ODataBatch.prototype, "batchId", {
	        get: function () {
	            return this._batchId;
	        },
	        enumerable: true,
	        configurable: true
	    });
	    /**
	     * Adds a request to a batch (not designed for public use)
	     *
	     * @param url The full url of the request
	     * @param method The http method GET, POST, etc
	     * @param options Any options to include in the request
	     * @param parser The parser that will hadle the results of the request
	     */
	    ODataBatch.prototype.add = function (url, method, options, parser) {
	        var info = {
	            method: method.toUpperCase(),
	            options: options,
	            parser: parser,
	            reject: null,
	            resolve: null,
	            url: url,
	        };
	        var p = new Promise(function (resolve, reject) {
	            info.resolve = resolve;
	            info.reject = reject;
	        });
	        this._requests.push(info);
	        return p;
	    };
	    /**
	     * Adds a dependency insuring that some set of actions will occur before a batch is processed.
	     * MUST be cleared using the returned resolve delegate to allow batches to run
	     */
	    ODataBatch.prototype.addDependency = function () {
	        var resolver;
	        var promise = new Promise(function (resolve) {
	            resolver = resolve;
	        });
	        this._dependencies.push(promise);
	        return resolver;
	    };
	    /**
	     * Execute the current batch and resolve the associated promises
	     *
	     * @returns A promise which will be resolved once all of the batch's child promises have resolved
	     */
	    ODataBatch.prototype.execute = function () {
	        var _this = this;
	        // we need to check the dependencies twice due to how different engines handle things.
	        // We can get a second set of promises added after the first set resolve
	        return Promise.all(this._dependencies).then(function () { return Promise.all(_this._dependencies); }).then(function () { return _this.executeImpl(); });
	    };
	    ODataBatch.prototype.executeImpl = function () {
	        var _this = this;
	        logging_1.Logger.write("[" + this.batchId + "] (" + (new Date()).getTime() + ") Executing batch with " + this._requests.length + " requests.", logging_1.LogLevel.Info);
	        // if we don't have any requests, don't bother sending anything
	        // this could be due to caching further upstream, or just an empty batch
	        if (this._requests.length < 1) {
	            logging_1.Logger.write("Resolving empty batch.", logging_1.LogLevel.Info);
	            return Promise.resolve();
	        }
	        // creating the client here allows the url to be populated for nodejs client as well as potentially
	        // any other hacks needed for other types of clients. Essentially allows the absoluteRequestUrl
	        // below to be correct
	        var client = new httpclient_1.HttpClient();
	        // due to timing we need to get the absolute url here so we can use it for all the individual requests
	        // and for sending the entire batch
	        return util_1.Util.toAbsoluteUrl(this.baseUrl).then(function (absoluteRequestUrl) {
	            // build all the requests, send them, pipe results in order to parsers
	            var batchBody = [];
	            var currentChangeSetId = "";
	            for (var i = 0; i < _this._requests.length; i++) {
	                var reqInfo = _this._requests[i];
	                if (reqInfo.method === "GET") {
	                    if (currentChangeSetId.length > 0) {
	                        // end an existing change set
	                        batchBody.push("--changeset_" + currentChangeSetId + "--\n\n");
	                        currentChangeSetId = "";
	                    }
	                    batchBody.push("--batch_" + _this._batchId + "\n");
	                }
	                else {
	                    if (currentChangeSetId.length < 1) {
	                        // start new change set
	                        currentChangeSetId = util_1.Util.getGUID();
	                        batchBody.push("--batch_" + _this._batchId + "\n");
	                        batchBody.push("Content-Type: multipart/mixed; boundary=\"changeset_" + currentChangeSetId + "\"\n\n");
	                    }
	                    batchBody.push("--changeset_" + currentChangeSetId + "\n");
	                }
	                // common batch part prefix
	                batchBody.push("Content-Type: application/http\n");
	                batchBody.push("Content-Transfer-Encoding: binary\n\n");
	                var headers = {
	                    "Accept": "application/json;",
	                };
	                // this is the url of the individual request within the batch
	                var url = util_1.Util.isUrlAbsolute(reqInfo.url) ? reqInfo.url : util_1.Util.combinePaths(absoluteRequestUrl, reqInfo.url);
	                logging_1.Logger.write("[" + _this.batchId + "] (" + (new Date()).getTime() + ") Adding request " + reqInfo.method + " " + url + " to batch.", logging_1.LogLevel.Verbose);
	                if (reqInfo.method !== "GET") {
	                    var method = reqInfo.method;
	                    if (reqInfo.hasOwnProperty("options") && reqInfo.options.hasOwnProperty("headers") && typeof reqInfo.options.headers["X-HTTP-Method"] !== "undefined") {
	                        method = reqInfo.options.headers["X-HTTP-Method"];
	                        delete reqInfo.options.headers["X-HTTP-Method"];
	                    }
	                    batchBody.push(method + " " + url + " HTTP/1.1\n");
	                    headers = util_1.Util.extend(headers, { "Content-Type": "application/json;odata=verbose;charset=utf-8" });
	                }
	                else {
	                    batchBody.push(reqInfo.method + " " + url + " HTTP/1.1\n");
	                }
	                if (typeof pnplibconfig_1.RuntimeConfig.headers !== "undefined") {
	                    headers = util_1.Util.extend(headers, pnplibconfig_1.RuntimeConfig.headers);
	                }
	                if (reqInfo.options && reqInfo.options.headers) {
	                    headers = util_1.Util.extend(headers, reqInfo.options.headers);
	                }
	                for (var name_1 in headers) {
	                    if (headers.hasOwnProperty(name_1)) {
	                        batchBody.push(name_1 + ": " + headers[name_1] + "\n");
	                    }
	                }
	                batchBody.push("\n");
	                if (reqInfo.options.body) {
	                    batchBody.push(reqInfo.options.body + "\n\n");
	                }
	            }
	            if (currentChangeSetId.length > 0) {
	                // Close the changeset
	                batchBody.push("--changeset_" + currentChangeSetId + "--\n\n");
	                currentChangeSetId = "";
	            }
	            batchBody.push("--batch_" + _this._batchId + "--\n");
	            var batchHeaders = {
	                "Content-Type": "multipart/mixed; boundary=batch_" + _this._batchId,
	            };
	            var batchOptions = {
	                "body": batchBody.join(""),
	                "headers": batchHeaders,
	            };
	            logging_1.Logger.write("[" + _this.batchId + "] (" + (new Date()).getTime() + ") Sending batch request.", logging_1.LogLevel.Info);
	            return client.post(util_1.Util.combinePaths(absoluteRequestUrl, "/_api/$batch"), batchOptions)
	                .then(function (r) { return r.text(); })
	                .then(_this._parseResponse)
	                .then(function (responses) {
	                if (responses.length !== _this._requests.length) {
	                    throw new exceptions_1.BatchParseException("Could not properly parse responses to match requests in batch.");
	                }
	                logging_1.Logger.write("[" + _this.batchId + "] (" + (new Date()).getTime() + ") Resolving batched requests.", logging_1.LogLevel.Info);
	                return responses.reduce(function (chain, response, index) {
	                    var request = _this._requests[index];
	                    logging_1.Logger.write("[" + _this.batchId + "] (" + (new Date()).getTime() + ") Resolving batched request " + request.method + " " + request.url + ".", logging_1.LogLevel.Verbose);
	                    return chain.then(function (_) { return request.parser.parse(response).then(request.resolve).catch(request.reject); });
	                }, Promise.resolve());
	            });
	        });
	    };
	    /**
	     * Parses the response from a batch request into an array of Response instances
	     *
	     * @param body Text body of the response from the batch request
	     */
	    ODataBatch.prototype._parseResponse = function (body) {
	        return new Promise(function (resolve, reject) {
	            var responses = [];
	            var header = "--batchresponse_";
	            // Ex. "HTTP/1.1 500 Internal Server Error"
	            var statusRegExp = new RegExp("^HTTP/[0-9.]+ +([0-9]+) +(.*)", "i");
	            var lines = body.split("\n");
	            var state = "batch";
	            var status;
	            var statusText;
	            for (var i = 0; i < lines.length; ++i) {
	                var line = lines[i];
	                switch (state) {
	                    case "batch":
	                        if (line.substr(0, header.length) === header) {
	                            state = "batchHeaders";
	                        }
	                        else {
	                            if (line.trim() !== "") {
	                                throw new exceptions_1.BatchParseException("Invalid response, line " + i);
	                            }
	                        }
	                        break;
	                    case "batchHeaders":
	                        if (line.trim() === "") {
	                            state = "status";
	                        }
	                        break;
	                    case "status":
	                        var parts = statusRegExp.exec(line);
	                        if (parts.length !== 3) {
	                            throw new exceptions_1.BatchParseException("Invalid status, line " + i);
	                        }
	                        status = parseInt(parts[1], 10);
	                        statusText = parts[2];
	                        state = "statusHeaders";
	                        break;
	                    case "statusHeaders":
	                        if (line.trim() === "") {
	                            state = "body";
	                        }
	                        break;
	                    case "body":
	                        responses.push((status === 204) ? new Response() : new Response(line, { status: status, statusText: statusText }));
	                        state = "batch";
	                        break;
	                }
	            }
	            if (state !== "status") {
	                reject(new exceptions_1.BatchParseException("Unexpected end of input"));
	            }
	            resolve(responses);
	        });
	    };
	    return ODataBatch;
	}());
	exports.ODataBatch = ODataBatch;
	var TextFileParser = (function () {
	    function TextFileParser() {
	    }
	    TextFileParser.prototype.parse = function (r) {
	        return r.text();
	    };
	    return TextFileParser;
	}());
	exports.TextFileParser = TextFileParser;
	var BlobFileParser = (function () {
	    function BlobFileParser() {
	    }
	    BlobFileParser.prototype.parse = function (r) {
	        return r.blob();
	    };
	    return BlobFileParser;
	}());
	exports.BlobFileParser = BlobFileParser;
	var JSONFileParser = (function () {
	    function JSONFileParser() {
	    }
	    JSONFileParser.prototype.parse = function (r) {
	        return r.json();
	    };
	    return JSONFileParser;
	}());
	exports.JSONFileParser = JSONFileParser;
	var BufferFileParser = (function () {
	    function BufferFileParser() {
	    }
	    BufferFileParser.prototype.parse = function (r) {
	        if (util_1.Util.isFunction(r.arrayBuffer)) {
	            return r.arrayBuffer();
	        }
	        return r.buffer();
	    };
	    return BufferFileParser;
	}());
	exports.BufferFileParser = BufferFileParser;


/***/ },
/* 24 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	Object.defineProperty(exports, "__esModule", { value: true });
	var digestcache_1 = __webpack_require__(25);
	var util_1 = __webpack_require__(12);
	var pnplibconfig_1 = __webpack_require__(15);
	var exceptions_1 = __webpack_require__(26);
	var HttpClient = (function () {
	    function HttpClient() {
	        this._impl = pnplibconfig_1.RuntimeConfig.fetchClientFactory();
	        this._digestCache = new digestcache_1.DigestCache(this);
	    }
	    HttpClient.prototype.fetch = function (url, options) {
	        var _this = this;
	        if (options === void 0) { options = {}; }
	        var opts = util_1.Util.extend(options, { cache: "no-cache", credentials: "same-origin" }, true);
	        var headers = new Headers();
	        // first we add the global headers so they can be overwritten by any passed in locally to this call
	        this.mergeHeaders(headers, pnplibconfig_1.RuntimeConfig.headers);
	        // second we add the local options so we can overwrite the globals
	        this.mergeHeaders(headers, options.headers);
	        // lastly we apply any default headers we need that may not exist
	        if (!headers.has("Accept")) {
	            headers.append("Accept", "application/json");
	        }
	        if (!headers.has("Content-Type")) {
	            headers.append("Content-Type", "application/json;odata=verbose;charset=utf-8");
	        }
	        if (!headers.has("X-ClientService-ClientTag")) {
	            headers.append("X-ClientService-ClientTag", "PnPCoreJS:2.0.4");
	        }
	        opts = util_1.Util.extend(opts, { headers: headers });
	        if (opts.method && opts.method.toUpperCase() !== "GET") {
	            if (!headers.has("X-RequestDigest")) {
	                var index = url.indexOf("_api/");
	                if (index < 0) {
	                    throw new exceptions_1.APIUrlException();
	                }
	                var webUrl = url.substr(0, index);
	                return this._digestCache.getDigest(webUrl)
	                    .then(function (digest) {
	                    headers.append("X-RequestDigest", digest);
	                    return _this.fetchRaw(url, opts);
	                });
	            }
	        }
	        return this.fetchRaw(url, opts);
	    };
	    HttpClient.prototype.fetchRaw = function (url, options) {
	        var _this = this;
	        if (options === void 0) { options = {}; }
	        // here we need to normalize the headers
	        var rawHeaders = new Headers();
	        this.mergeHeaders(rawHeaders, options.headers);
	        options = util_1.Util.extend(options, { headers: rawHeaders });
	        var retry = function (ctx) {
	            _this._impl.fetch(url, options).then(function (response) { return ctx.resolve(response); }).catch(function (response) {
	                // grab our current delay
	                var delay = ctx.delay;
	                // Check if request was throttled - http status code 429
	                // Check is request failed due to server unavailable - http status code 503
	                if (response.status !== 429 && response.status !== 503) {
	                    ctx.reject(response);
	                }
	                // Increment our counters.
	                ctx.delay *= 2;
	                ctx.attempts++;
	                // If we have exceeded the retry count, reject.
	                if (ctx.retryCount <= ctx.attempts) {
	                    ctx.reject(response);
	                }
	                // Set our retry timeout for {delay} milliseconds.
	                setTimeout(util_1.Util.getCtxCallback(_this, retry, ctx), delay);
	            });
	        };
	        return new Promise(function (resolve, reject) {
	            var retryContext = {
	                attempts: 0,
	                delay: 100,
	                reject: reject,
	                resolve: resolve,
	                retryCount: 7,
	            };
	            retry.call(_this, retryContext);
	        });
	    };
	    HttpClient.prototype.get = function (url, options) {
	        if (options === void 0) { options = {}; }
	        var opts = util_1.Util.extend(options, { method: "GET" });
	        return this.fetch(url, opts);
	    };
	    HttpClient.prototype.post = function (url, options) {
	        if (options === void 0) { options = {}; }
	        var opts = util_1.Util.extend(options, { method: "POST" });
	        return this.fetch(url, opts);
	    };
	    HttpClient.prototype.patch = function (url, options) {
	        if (options === void 0) { options = {}; }
	        var opts = util_1.Util.extend(options, { method: "PATCH" });
	        return this.fetch(url, opts);
	    };
	    HttpClient.prototype.delete = function (url, options) {
	        if (options === void 0) { options = {}; }
	        var opts = util_1.Util.extend(options, { method: "DELETE" });
	        return this.fetch(url, opts);
	    };
	    HttpClient.prototype.mergeHeaders = function (target, source) {
	        if (typeof source !== "undefined" && source !== null) {
	            var temp = new Request("", { headers: source });
	            temp.headers.forEach(function (value, name) {
	                target.append(name, value);
	            });
	        }
	    };
	    return HttpClient;
	}());
	exports.HttpClient = HttpClient;
	;


/***/ },
/* 25 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	Object.defineProperty(exports, "__esModule", { value: true });
	var collections_1 = __webpack_require__(18);
	var util_1 = __webpack_require__(12);
	var odata_1 = __webpack_require__(23);
	var CachedDigest = (function () {
	    function CachedDigest() {
	    }
	    return CachedDigest;
	}());
	exports.CachedDigest = CachedDigest;
	// allows for the caching of digests across all HttpClient's which each have their own DigestCache wrapper.
	var digests = new collections_1.Dictionary();
	var DigestCache = (function () {
	    function DigestCache(_httpClient, _digests) {
	        if (_digests === void 0) { _digests = digests; }
	        this._httpClient = _httpClient;
	        this._digests = _digests;
	    }
	    DigestCache.prototype.getDigest = function (webUrl) {
	        var _this = this;
	        var cachedDigest = this._digests.get(webUrl);
	        if (cachedDigest !== null) {
	            var now = new Date();
	            if (now < cachedDigest.expiration) {
	                return Promise.resolve(cachedDigest.value);
	            }
	        }
	        var url = util_1.Util.combinePaths(webUrl, "/_api/contextinfo");
	        return this._httpClient.fetchRaw(url, {
	            cache: "no-cache",
	            credentials: "same-origin",
	            headers: {
	                "Accept": "application/json;odata=verbose",
	                "Content-type": "application/json;odata=verbose;charset=utf-8",
	            },
	            method: "POST",
	        }).then(function (response) {
	            var parser = new odata_1.ODataDefaultParser();
	            return parser.parse(response).then(function (d) { return d.GetContextWebInformation; });
	        }).then(function (data) {
	            var newCachedDigest = new CachedDigest();
	            newCachedDigest.value = data.FormDigestValue;
	            var seconds = data.FormDigestTimeoutSeconds;
	            var expiration = new Date();
	            expiration.setTime(expiration.getTime() + 1000 * seconds);
	            newCachedDigest.expiration = expiration;
	            _this._digests.add(webUrl, newCachedDigest);
	            return newCachedDigest.value;
	        });
	    };
	    DigestCache.prototype.clear = function () {
	        this._digests.clear();
	    };
	    return DigestCache;
	}());
	exports.DigestCache = DigestCache;


/***/ },
/* 26 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	var __extends = (this && this.__extends) || (function () {
	    var extendStatics = Object.setPrototypeOf ||
	        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
	        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
	    return function (d, b) {
	        extendStatics(d, b);
	        function __() { this.constructor = d; }
	        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
	    };
	})();
	Object.defineProperty(exports, "__esModule", { value: true });
	var logging_1 = __webpack_require__(14);
	function defaultLog(error) {
	    logging_1.Logger.log({ data: {}, level: logging_1.LogLevel.Error, message: "[" + error.name + "]::" + error.message });
	}
	/**
	 * Represents an exception with an HttpClient request
	 *
	 */
	var ProcessHttpClientResponseException = (function (_super) {
	    __extends(ProcessHttpClientResponseException, _super);
	    function ProcessHttpClientResponseException(status, statusText, data) {
	        var _this = _super.call(this, "Error making HttpClient request in queryable: [" + status + "] " + statusText) || this;
	        _this.status = status;
	        _this.statusText = statusText;
	        _this.data = data;
	        _this.name = "ProcessHttpClientResponseException";
	        logging_1.Logger.log({ data: _this.data, level: logging_1.LogLevel.Error, message: _this.message });
	        return _this;
	    }
	    return ProcessHttpClientResponseException;
	}(Error));
	exports.ProcessHttpClientResponseException = ProcessHttpClientResponseException;
	var NoCacheAvailableException = (function (_super) {
	    __extends(NoCacheAvailableException, _super);
	    function NoCacheAvailableException(msg) {
	        if (msg === void 0) { msg = "Cannot create a caching configuration provider since cache is not available."; }
	        var _this = _super.call(this, msg) || this;
	        _this.name = "NoCacheAvailableException";
	        defaultLog(_this);
	        return _this;
	    }
	    return NoCacheAvailableException;
	}(Error));
	exports.NoCacheAvailableException = NoCacheAvailableException;
	var APIUrlException = (function (_super) {
	    __extends(APIUrlException, _super);
	    function APIUrlException(msg) {
	        if (msg === void 0) { msg = "Unable to determine API url."; }
	        var _this = _super.call(this, msg) || this;
	        _this.name = "APIUrlException";
	        defaultLog(_this);
	        return _this;
	    }
	    return APIUrlException;
	}(Error));
	exports.APIUrlException = APIUrlException;
	var AuthUrlException = (function (_super) {
	    __extends(AuthUrlException, _super);
	    function AuthUrlException(data, msg) {
	        if (msg === void 0) { msg = "Auth URL Endpoint could not be determined from data. Data logged."; }
	        var _this = _super.call(this, msg) || this;
	        _this.name = "APIUrlException";
	        logging_1.Logger.log({ data: data, level: logging_1.LogLevel.Error, message: _this.message });
	        return _this;
	    }
	    return AuthUrlException;
	}(Error));
	exports.AuthUrlException = AuthUrlException;
	var NodeFetchClientUnsupportedException = (function (_super) {
	    __extends(NodeFetchClientUnsupportedException, _super);
	    function NodeFetchClientUnsupportedException(msg) {
	        if (msg === void 0) { msg = "Using NodeFetchClient in the browser is not supported."; }
	        var _this = _super.call(this, msg) || this;
	        _this.name = "NodeFetchClientUnsupportedException";
	        defaultLog(_this);
	        return _this;
	    }
	    return NodeFetchClientUnsupportedException;
	}(Error));
	exports.NodeFetchClientUnsupportedException = NodeFetchClientUnsupportedException;
	var SPRequestExecutorUndefinedException = (function (_super) {
	    __extends(SPRequestExecutorUndefinedException, _super);
	    function SPRequestExecutorUndefinedException() {
	        var _this = this;
	        var msg = [
	            "SP.RequestExecutor is undefined. ",
	            "Load the SP.RequestExecutor.js library (/_layouts/15/SP.RequestExecutor.js) before loading the PnP JS Core library.",
	        ].join(" ");
	        _this = _super.call(this, msg) || this;
	        _this.name = "SPRequestExecutorUndefinedException";
	        defaultLog(_this);
	        return _this;
	    }
	    return SPRequestExecutorUndefinedException;
	}(Error));
	exports.SPRequestExecutorUndefinedException = SPRequestExecutorUndefinedException;
	var MaxCommentLengthException = (function (_super) {
	    __extends(MaxCommentLengthException, _super);
	    function MaxCommentLengthException(msg) {
	        if (msg === void 0) { msg = "The maximum comment length is 1023 characters."; }
	        var _this = _super.call(this, msg) || this;
	        _this.name = "MaxCommentLengthException";
	        defaultLog(_this);
	        return _this;
	    }
	    return MaxCommentLengthException;
	}(Error));
	exports.MaxCommentLengthException = MaxCommentLengthException;
	var NotSupportedInBatchException = (function (_super) {
	    __extends(NotSupportedInBatchException, _super);
	    function NotSupportedInBatchException(operation) {
	        if (operation === void 0) { operation = "This operation"; }
	        var _this = _super.call(this, operation + " is not supported as part of a batch.") || this;
	        _this.name = "NotSupportedInBatchException";
	        defaultLog(_this);
	        return _this;
	    }
	    return NotSupportedInBatchException;
	}(Error));
	exports.NotSupportedInBatchException = NotSupportedInBatchException;
	var ODataIdException = (function (_super) {
	    __extends(ODataIdException, _super);
	    function ODataIdException(data, msg) {
	        if (msg === void 0) { msg = "Could not extract odata id in object, you may be using nometadata. Object data logged to logger."; }
	        var _this = _super.call(this, msg) || this;
	        _this.name = "ODataIdException";
	        logging_1.Logger.log({ data: data, level: logging_1.LogLevel.Error, message: _this.message });
	        return _this;
	    }
	    return ODataIdException;
	}(Error));
	exports.ODataIdException = ODataIdException;
	var BatchParseException = (function (_super) {
	    __extends(BatchParseException, _super);
	    function BatchParseException(msg) {
	        var _this = _super.call(this, msg) || this;
	        _this.name = "BatchParseException";
	        defaultLog(_this);
	        return _this;
	    }
	    return BatchParseException;
	}(Error));
	exports.BatchParseException = BatchParseException;
	var AlreadyInBatchException = (function (_super) {
	    __extends(AlreadyInBatchException, _super);
	    function AlreadyInBatchException(msg) {
	        if (msg === void 0) { msg = "This query is already part of a batch."; }
	        var _this = _super.call(this, msg) || this;
	        _this.name = "AlreadyInBatchException";
	        defaultLog(_this);
	        return _this;
	    }
	    return AlreadyInBatchException;
	}(Error));
	exports.AlreadyInBatchException = AlreadyInBatchException;
	var FunctionExpectedException = (function (_super) {
	    __extends(FunctionExpectedException, _super);
	    function FunctionExpectedException(msg) {
	        if (msg === void 0) { msg = "This query is already part of a batch."; }
	        var _this = _super.call(this, msg) || this;
	        _this.name = "FunctionExpectedException";
	        defaultLog(_this);
	        return _this;
	    }
	    return FunctionExpectedException;
	}(Error));
	exports.FunctionExpectedException = FunctionExpectedException;
	var UrlException = (function (_super) {
	    __extends(UrlException, _super);
	    function UrlException(msg) {
	        var _this = _super.call(this, msg) || this;
	        _this.name = "UrlException";
	        defaultLog(_this);
	        return _this;
	    }
	    return UrlException;
	}(Error));
	exports.UrlException = UrlException;


/***/ },
/* 27 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
	    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
	    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
	    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
	    return c > 3 && r && Object.defineProperty(target, key, r), r;
	};
	Object.defineProperty(exports, "__esModule", { value: true });
	var caching_1 = __webpack_require__(28);
	var httpclient_1 = __webpack_require__(24);
	var logging_1 = __webpack_require__(14);
	var util_1 = __webpack_require__(12);
	/**
	 * Resolves the context's result value
	 *
	 * @param context The current context
	 */
	function returnResult(context) {
	    logging_1.Logger.log({
	        data: context.result,
	        level: logging_1.LogLevel.Verbose,
	        message: "[" + context.requestId + "] (" + (new Date()).getTime() + ") Returning result, see data property for value.",
	    });
	    return Promise.resolve(context.result);
	}
	/**
	 * Sets the result on the context
	 */
	function setResult(context, value) {
	    return new Promise(function (resolve) {
	        context.result = value;
	        context.hasResult = true;
	        resolve(context);
	    });
	}
	exports.setResult = setResult;
	/**
	 * Invokes the next method in the provided context's pipeline
	 *
	 * @param c The current request context
	 */
	function next(c) {
	    if (c.pipeline.length < 1) {
	        return Promise.resolve(c);
	    }
	    return c.pipeline.shift()(c);
	}
	/**
	 * Executes the current request context's pipeline
	 *
	 * @param context Current context
	 */
	function pipe(context) {
	    return next(context)
	        .then(function (ctx) { return returnResult(ctx); })
	        .catch(function (e) {
	        logging_1.Logger.log({
	            data: e,
	            level: logging_1.LogLevel.Error,
	            message: "Error in request pipeline: " + e.message,
	        });
	        throw e;
	    });
	}
	exports.pipe = pipe;
	/**
	 * decorator factory applied to methods in the pipeline to control behavior
	 */
	function requestPipelineMethod(alwaysRun) {
	    if (alwaysRun === void 0) { alwaysRun = false; }
	    return function (target, propertyKey, descriptor) {
	        var method = descriptor.value;
	        descriptor.value = function () {
	            var args = [];
	            for (var _i = 0; _i < arguments.length; _i++) {
	                args[_i] = arguments[_i];
	            }
	            // if we have a result already in the pipeline, pass it along and don't call the tagged method
	            if (!alwaysRun && args.length > 0 && args[0].hasOwnProperty("hasResult") && args[0].hasResult) {
	                logging_1.Logger.write("[" + args[0].requestId + "] (" + (new Date()).getTime() + ") Skipping request pipeline method " + propertyKey + ", existing result in pipeline.", logging_1.LogLevel.Verbose);
	                return Promise.resolve(args[0]);
	            }
	            // apply the tagged method
	            logging_1.Logger.write("[" + args[0].requestId + "] (" + (new Date()).getTime() + ") Calling request pipeline method " + propertyKey + ".", logging_1.LogLevel.Verbose);
	            // then chain the next method in the context's pipeline - allows for dynamic pipeline
	            return method.apply(target, args).then(function (ctx) { return next(ctx); });
	        };
	    };
	}
	exports.requestPipelineMethod = requestPipelineMethod;
	/**
	 * Contains the methods used within the request pipeline
	 */
	var PipelineMethods = (function () {
	    function PipelineMethods() {
	    }
	    /**
	     * Logs the start of the request
	     */
	    PipelineMethods.logStart = function (context) {
	        return new Promise(function (resolve) {
	            logging_1.Logger.log({
	                data: logging_1.Logger.activeLogLevel === logging_1.LogLevel.Info ? {} : context,
	                level: logging_1.LogLevel.Info,
	                message: "[" + context.requestId + "] (" + (new Date()).getTime() + ") Beginning " + context.verb + " request (" + context.requestAbsoluteUrl + ")",
	            });
	            resolve(context);
	        });
	    };
	    /**
	     * Handles caching of the request
	     */
	    PipelineMethods.caching = function (context) {
	        return new Promise(function (resolve) {
	            // handle caching, if applicable
	            if (context.verb === "GET" && context.isCached) {
	                logging_1.Logger.write("[" + context.requestId + "] (" + (new Date()).getTime() + ") Caching is enabled for request, checking cache...", logging_1.LogLevel.Info);
	                var cacheOptions = new caching_1.CachingOptions(context.requestAbsoluteUrl.toLowerCase());
	                if (typeof context.cachingOptions !== "undefined") {
	                    cacheOptions = util_1.Util.extend(cacheOptions, context.cachingOptions);
	                }
	                // we may not have a valid store
	                if (cacheOptions.store !== null) {
	                    // check if we have the data in cache and if so resolve the promise and return
	                    var data = cacheOptions.store.get(cacheOptions.key);
	                    if (data !== null) {
	                        // ensure we clear any help batch dependency we are resolving from the cache
	                        logging_1.Logger.log({
	                            data: logging_1.Logger.activeLogLevel === logging_1.LogLevel.Info ? {} : data,
	                            level: logging_1.LogLevel.Info,
	                            message: "[" + context.requestId + "] (" + (new Date()).getTime() + ") Value returned from cache.",
	                        });
	                        context.batchDependency();
	                        return setResult(context, data).then(function (ctx) { return resolve(ctx); });
	                    }
	                }
	                logging_1.Logger.write("[" + context.requestId + "] (" + (new Date()).getTime() + ") Value not found in cache.", logging_1.LogLevel.Info);
	                // if we don't then wrap the supplied parser in the caching parser wrapper
	                // and send things on their way
	                context.parser = new caching_1.CachingParserWrapper(context.parser, cacheOptions);
	            }
	            return resolve(context);
	        });
	    };
	    /**
	     * Sends the request
	     */
	    PipelineMethods.send = function (context) {
	        return new Promise(function (resolve, reject) {
	            // send or batch the request
	            if (context.isBatched) {
	                // we are in a batch, so add to batch, remove dependency, and resolve with the batch's promise
	                var p = context.batch.add(context.requestAbsoluteUrl, context.verb, context.options, context.parser);
	                // we release the dependency here to ensure the batch does not execute until the request is added to the batch
	                context.batchDependency();
	                logging_1.Logger.write("[" + context.requestId + "] (" + (new Date()).getTime() + ") Batching request in batch " + context.batch.batchId + ".", logging_1.LogLevel.Info);
	                // we set the result as the promise which will be resolved by the batch's execution
	                resolve(setResult(context, p));
	            }
	            else {
	                logging_1.Logger.write("[" + context.requestId + "] (" + (new Date()).getTime() + ") Sending request.", logging_1.LogLevel.Info);
	                // we are not part of a batch, so proceed as normal
	                var client = new httpclient_1.HttpClient();
	                var opts = util_1.Util.extend(context.options || {}, { method: context.verb });
	                client.fetch(context.requestAbsoluteUrl, opts)
	                    .then(function (response) { return context.parser.parse(response); })
	                    .then(function (result) { return setResult(context, result); })
	                    .then(function (ctx) { return resolve(ctx); })
	                    .catch(function (e) { return reject(e); });
	            }
	        });
	    };
	    /**
	     * Logs the end of the request
	     */
	    PipelineMethods.logEnd = function (context) {
	        return new Promise(function (resolve) {
	            if (context.isBatched) {
	                logging_1.Logger.log({
	                    data: logging_1.Logger.activeLogLevel === logging_1.LogLevel.Info ? {} : context,
	                    level: logging_1.LogLevel.Info,
	                    message: "[" + context.requestId + "] (" + (new Date()).getTime() + ") " + context.verb + " request will complete in batch " + context.batch.batchId + ".",
	                });
	            }
	            else {
	                logging_1.Logger.log({
	                    data: logging_1.Logger.activeLogLevel === logging_1.LogLevel.Info ? {} : context,
	                    level: logging_1.LogLevel.Info,
	                    message: "[" + context.requestId + "] (" + (new Date()).getTime() + ") Completing " + context.verb + " request.",
	                });
	            }
	            resolve(context);
	        });
	    };
	    Object.defineProperty(PipelineMethods, "default", {
	        get: function () {
	            return [
	                PipelineMethods.logStart,
	                PipelineMethods.caching,
	                PipelineMethods.send,
	                PipelineMethods.logEnd,
	            ];
	        },
	        enumerable: true,
	        configurable: true
	    });
	    return PipelineMethods;
	}());
	__decorate([
	    requestPipelineMethod(true)
	], PipelineMethods, "logStart", null);
	__decorate([
	    requestPipelineMethod()
	], PipelineMethods, "caching", null);
	__decorate([
	    requestPipelineMethod()
	], PipelineMethods, "send", null);
	__decorate([
	    requestPipelineMethod(true)
	], PipelineMethods, "logEnd", null);
	exports.PipelineMethods = PipelineMethods;


/***/ },
/* 28 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	Object.defineProperty(exports, "__esModule", { value: true });
	var storage_1 = __webpack_require__(17);
	var util_1 = __webpack_require__(12);
	var pnplibconfig_1 = __webpack_require__(15);
	var CachingOptions = (function () {
	    function CachingOptions(key) {
	        this.key = key;
	        this.expiration = util_1.Util.dateAdd(new Date(), "second", pnplibconfig_1.RuntimeConfig.defaultCachingTimeoutSeconds);
	        this.storeName = pnplibconfig_1.RuntimeConfig.defaultCachingStore;
	    }
	    Object.defineProperty(CachingOptions.prototype, "store", {
	        get: function () {
	            if (this.storeName === "local") {
	                return CachingOptions.storage.local;
	            }
	            else {
	                return CachingOptions.storage.session;
	            }
	        },
	        enumerable: true,
	        configurable: true
	    });
	    return CachingOptions;
	}());
	CachingOptions.storage = new storage_1.PnPClientStorage();
	exports.CachingOptions = CachingOptions;
	var CachingParserWrapper = (function () {
	    function CachingParserWrapper(_parser, _cacheOptions) {
	        this._parser = _parser;
	        this._cacheOptions = _cacheOptions;
	    }
	    CachingParserWrapper.prototype.parse = function (response) {
	        var _this = this;
	        // add this to the cache based on the options
	        return this._parser.parse(response).then(function (data) {
	            if (_this._cacheOptions.store !== null) {
	                _this._cacheOptions.store.put(_this._cacheOptions.key, data, _this._cacheOptions.expiration);
	            }
	            return data;
	        });
	    };
	    return CachingParserWrapper;
	}());
	exports.CachingParserWrapper = CachingParserWrapper;


/***/ },
/* 29 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	var __extends = (this && this.__extends) || (function () {
	    var extendStatics = Object.setPrototypeOf ||
	        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
	        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
	    return function (d, b) {
	        extendStatics(d, b);
	        function __() { this.constructor = d; }
	        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
	    };
	})();
	Object.defineProperty(exports, "__esModule", { value: true });
	var queryable_1 = __webpack_require__(22);
	var SearchSuggest = (function (_super) {
	    __extends(SearchSuggest, _super);
	    function SearchSuggest(baseUrl, path) {
	        if (path === void 0) { path = "_api/search/suggest"; }
	        return _super.call(this, baseUrl, path) || this;
	    }
	    SearchSuggest.prototype.execute = function (query) {
	        this.mapQueryToQueryString(query);
	        return this.get().then(function (response) { return new SearchSuggestResult(response); });
	    };
	    SearchSuggest.prototype.mapQueryToQueryString = function (query) {
	        this.query.add("querytext", "'" + query.querytext + "'");
	        if (query.hasOwnProperty("count")) {
	            this.query.add("inumberofquerysuggestions", query.count.toString());
	        }
	        if (query.hasOwnProperty("personalCount")) {
	            this.query.add("inumberofresultsuggestions", query.personalCount.toString());
	        }
	        if (query.hasOwnProperty("preQuery")) {
	            this.query.add("fprequerysuggestions", query.preQuery.toString());
	        }
	        if (query.hasOwnProperty("hitHighlighting")) {
	            this.query.add("fhithighlighting", query.hitHighlighting.toString());
	        }
	        if (query.hasOwnProperty("capitalize")) {
	            this.query.add("fcapitalizefirstletters", query.capitalize.toString());
	        }
	        if (query.hasOwnProperty("culture")) {
	            this.query.add("culture", query.culture.toString());
	        }
	        if (query.hasOwnProperty("stemming")) {
	            this.query.add("enablestemming", query.stemming.toString());
	        }
	        if (query.hasOwnProperty("includePeople")) {
	            this.query.add("showpeoplenamesuggestions", query.includePeople.toString());
	        }
	        if (query.hasOwnProperty("queryRules")) {
	            this.query.add("enablequeryrules", query.queryRules.toString());
	        }
	        if (query.hasOwnProperty("prefixMatch")) {
	            this.query.add("fprefixmatchallterms", query.prefixMatch.toString());
	        }
	    };
	    return SearchSuggest;
	}(queryable_1.QueryableInstance));
	exports.SearchSuggest = SearchSuggest;
	var SearchSuggestResult = (function () {
	    function SearchSuggestResult(json) {
	        if (json.hasOwnProperty("suggest")) {
	            // verbose
	            this.PeopleNames = json.suggest.PeopleNames.results;
	            this.PersonalResults = json.suggest.PersonalResults.results;
	            this.Queries = json.suggest.Queries.results;
	        }
	        else {
	            this.PeopleNames = json.PeopleNames;
	            this.PersonalResults = json.PersonalResults;
	            this.Queries = json.Queries;
	        }
	    }
	    return SearchSuggestResult;
	}());
	exports.SearchSuggestResult = SearchSuggestResult;


/***/ },
/* 30 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	var __extends = (this && this.__extends) || (function () {
	    var extendStatics = Object.setPrototypeOf ||
	        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
	        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
	    return function (d, b) {
	        extendStatics(d, b);
	        function __() { this.constructor = d; }
	        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
	    };
	})();
	Object.defineProperty(exports, "__esModule", { value: true });
	var queryable_1 = __webpack_require__(22);
	var webs_1 = __webpack_require__(31);
	var usercustomactions_1 = __webpack_require__(49);
	var odata_1 = __webpack_require__(23);
	var features_1 = __webpack_require__(51);
	/**
	 * Describes a site collection
	 *
	 */
	var Site = (function (_super) {
	    __extends(Site, _super);
	    /**
	     * Creates a new instance of the RoleAssignments class
	     *
	     * @param baseUrl The url or Queryable which forms the parent of this fields collection
	     */
	    function Site(baseUrl, path) {
	        if (path === void 0) { path = "_api/site"; }
	        return _super.call(this, baseUrl, path) || this;
	    }
	    Object.defineProperty(Site.prototype, "rootWeb", {
	        /**
	         * Gets the root web of the site collection
	         *
	         */
	        get: function () {
	            return new webs_1.Web(this, "rootweb");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Site.prototype, "features", {
	        /**
	         * Gets the active features for this site
	         *
	         */
	        get: function () {
	            return new features_1.Features(this);
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Site.prototype, "userCustomActions", {
	        /**
	         * Get all custom actions on a site collection
	         *
	         */
	        get: function () {
	            return new usercustomactions_1.UserCustomActions(this);
	        },
	        enumerable: true,
	        configurable: true
	    });
	    /**
	     * Gets the context information for the site.
	     */
	    Site.prototype.getContextInfo = function () {
	        var q = new Site(this.parentUrl, "_api/contextinfo");
	        return q.post().then(function (data) {
	            if (data.hasOwnProperty("GetContextWebInformation")) {
	                var info = data.GetContextWebInformation;
	                info.SupportedSchemaVersions = info.SupportedSchemaVersions.results;
	                return info;
	            }
	            else {
	                return data;
	            }
	        });
	    };
	    /**
	     * Gets the document libraries on a site. Static method. (SharePoint Online only)
	     *
	     * @param absoluteWebUrl The absolute url of the web whose document libraries should be returned
	     */
	    Site.prototype.getDocumentLibraries = function (absoluteWebUrl) {
	        var q = new queryable_1.Queryable("", "_api/sp.web.getdocumentlibraries(@v)");
	        q.query.add("@v", "'" + absoluteWebUrl + "'");
	        return q.get().then(function (data) {
	            if (data.hasOwnProperty("GetDocumentLibraries")) {
	                return data.GetDocumentLibraries;
	            }
	            else {
	                return data;
	            }
	        });
	    };
	    /**
	     * Gets the site URL from a page URL.
	     *
	     * @param absolutePageUrl The absolute url of the page
	     */
	    Site.prototype.getWebUrlFromPageUrl = function (absolutePageUrl) {
	        var q = new queryable_1.Queryable("", "_api/sp.web.getweburlfrompageurl(@v)");
	        q.query.add("@v", "'" + absolutePageUrl + "'");
	        return q.get().then(function (data) {
	            if (data.hasOwnProperty("GetWebUrlFromPageUrl")) {
	                return data.GetWebUrlFromPageUrl;
	            }
	            else {
	                return data;
	            }
	        });
	    };
	    /**
	     * Creates a new batch for requests within the context of context this site
	     *
	     */
	    Site.prototype.createBatch = function () {
	        return new odata_1.ODataBatch(this.parentUrl);
	    };
	    return Site;
	}(queryable_1.QueryableInstance));
	exports.Site = Site;


/***/ },
/* 31 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	var __extends = (this && this.__extends) || (function () {
	    var extendStatics = Object.setPrototypeOf ||
	        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
	        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
	    return function (d, b) {
	        extendStatics(d, b);
	        function __() { this.constructor = d; }
	        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
	    };
	})();
	var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
	    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
	    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
	    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
	    return c > 3 && r && Object.defineProperty(target, key, r), r;
	};
	Object.defineProperty(exports, "__esModule", { value: true });
	var queryable_1 = __webpack_require__(22);
	var lists_1 = __webpack_require__(32);
	var fields_1 = __webpack_require__(46);
	var navigation_1 = __webpack_require__(50);
	var sitegroups_1 = __webpack_require__(37);
	var contenttypes_1 = __webpack_require__(43);
	var folders_1 = __webpack_require__(40);
	var roles_1 = __webpack_require__(36);
	var files_1 = __webpack_require__(41);
	var util_1 = __webpack_require__(12);
	var lists_2 = __webpack_require__(32);
	var siteusers_1 = __webpack_require__(38);
	var usercustomactions_1 = __webpack_require__(49);
	var odata_1 = __webpack_require__(23);
	var features_1 = __webpack_require__(51);
	var decorators_1 = __webpack_require__(13);
	var queryableshareable_1 = __webpack_require__(34);
	var relateditems_1 = __webpack_require__(52);
	var Webs = (function (_super) {
	    __extends(Webs, _super);
	    function Webs(baseUrl, webPath) {
	        if (webPath === void 0) { webPath = "webs"; }
	        return _super.call(this, baseUrl, webPath) || this;
	    }
	    /**
	     * Adds a new web to the collection
	     *
	     * @param title The new web's title
	     * @param url The new web's relative url
	     * @param description The web web's description
	     * @param template The web's template
	     * @param language The language code to use for this web
	     * @param inheritPermissions If true permissions will be inherited from the partent web
	     * @param additionalSettings Will be passed as part of the web creation body
	     */
	    Webs.prototype.add = function (title, url, description, template, language, inheritPermissions, additionalSettings) {
	        if (description === void 0) { description = ""; }
	        if (template === void 0) { template = "STS"; }
	        if (language === void 0) { language = 1033; }
	        if (inheritPermissions === void 0) { inheritPermissions = true; }
	        if (additionalSettings === void 0) { additionalSettings = {}; }
	        var props = util_1.Util.extend({
	            Description: description,
	            Language: language,
	            Title: title,
	            Url: url,
	            UseSamePermissionsAsParentSite: inheritPermissions,
	            WebTemplate: template,
	        }, additionalSettings);
	        var postBody = JSON.stringify({
	            "parameters": util_1.Util.extend({
	                "__metadata": { "type": "SP.WebCreationInformation" },
	            }, props),
	        });
	        return this.clone(Webs, "add", true).post({ body: postBody }).then(function (data) {
	            return {
	                data: data,
	                web: new Web(odata_1.extractOdataId(data).replace(/_api\/web\/?/i, "")),
	            };
	        });
	    };
	    return Webs;
	}(queryable_1.QueryableCollection));
	exports.Webs = Webs;
	var WebInfos = (function (_super) {
	    __extends(WebInfos, _super);
	    function WebInfos(baseUrl, webPath) {
	        if (webPath === void 0) { webPath = "webinfos"; }
	        return _super.call(this, baseUrl, webPath) || this;
	    }
	    return WebInfos;
	}(queryable_1.QueryableCollection));
	exports.WebInfos = WebInfos;
	/**
	 * Describes a web
	 *
	 */
	var Web = (function (_super) {
	    __extends(Web, _super);
	    function Web(baseUrl, path) {
	        if (path === void 0) { path = "_api/web"; }
	        return _super.call(this, baseUrl, path) || this;
	    }
	    /**
	     * Creates a new web instance from the given url by indexing the location of the /_api/
	     * segment. If this is not found the method creates a new web with the entire string as
	     * supplied.
	     *
	     * @param url
	     */
	    Web.fromUrl = function (url, path) {
	        if (url === null) {
	            return new Web("");
	        }
	        var index = url.indexOf("_api/");
	        if (index > -1) {
	            return new Web(url.substr(0, index), path);
	        }
	        return new Web(url, path);
	    };
	    Object.defineProperty(Web.prototype, "webs", {
	        get: function () {
	            return new Webs(this);
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Web.prototype, "webinfos", {
	        get: function () {
	            return new WebInfos(this);
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Web.prototype, "contentTypes", {
	        /**
	         * Get the content types available in this web
	         *
	         */
	        get: function () {
	            return new contenttypes_1.ContentTypes(this);
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Web.prototype, "lists", {
	        /**
	         * Get the lists in this web
	         *
	         */
	        get: function () {
	            return new lists_1.Lists(this);
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Web.prototype, "fields", {
	        /**
	         * Gets the fields in this web
	         *
	         */
	        get: function () {
	            return new fields_1.Fields(this);
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Web.prototype, "features", {
	        /**
	         * Gets the active features for this web
	         *
	         */
	        get: function () {
	            return new features_1.Features(this);
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Web.prototype, "availablefields", {
	        /**
	         * Gets the available fields in this web
	         *
	         */
	        get: function () {
	            return new fields_1.Fields(this, "availablefields");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Web.prototype, "navigation", {
	        /**
	         * Get the navigation options in this web
	         *
	         */
	        get: function () {
	            return new navigation_1.Navigation(this);
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Web.prototype, "siteUsers", {
	        /**
	         * Gets the site users
	         *
	         */
	        get: function () {
	            return new siteusers_1.SiteUsers(this);
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Web.prototype, "siteGroups", {
	        /**
	         * Gets the site groups
	         *
	         */
	        get: function () {
	            return new sitegroups_1.SiteGroups(this);
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Web.prototype, "currentUser", {
	        /**
	         * Gets the current user
	         */
	        get: function () {
	            return new siteusers_1.CurrentUser(this);
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Web.prototype, "folders", {
	        /**
	         * Get the folders in this web
	         *
	         */
	        get: function () {
	            return new folders_1.Folders(this);
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Web.prototype, "userCustomActions", {
	        /**
	         * Get all custom actions on a site
	         *
	         */
	        get: function () {
	            return new usercustomactions_1.UserCustomActions(this);
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Web.prototype, "roleDefinitions", {
	        /**
	         * Gets the collection of RoleDefinition resources.
	         *
	         */
	        get: function () {
	            return new roles_1.RoleDefinitions(this);
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Web.prototype, "relatedItems", {
	        /**
	         * Provides an interface to manage related items
	         *
	         */
	        get: function () {
	            return relateditems_1.RelatedItemManagerImpl.FromUrl(this.toUrl());
	        },
	        enumerable: true,
	        configurable: true
	    });
	    /**
	     * Creates a new batch for requests within the context of context this web
	     *
	     */
	    Web.prototype.createBatch = function () {
	        return new odata_1.ODataBatch(this.parentUrl);
	    };
	    Object.defineProperty(Web.prototype, "rootFolder", {
	        /**
	         * The root folder of the web
	         */
	        get: function () {
	            return new folders_1.Folder(this, "rootFolder");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Web.prototype, "associatedOwnerGroup", {
	        get: function () {
	            return new sitegroups_1.SiteGroup(this, "associatedownergroup");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Web.prototype, "associatedMemberGroup", {
	        get: function () {
	            return new sitegroups_1.SiteGroup(this, "associatedmembergroup");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Web.prototype, "associatedVisitorGroup", {
	        get: function () {
	            return new sitegroups_1.SiteGroup(this, "associatedvisitorgroup");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    /**
	     * Get a folder by server relative url
	     *
	     * @param folderRelativeUrl the server relative path to the folder (including /sites/ if applicable)
	     */
	    Web.prototype.getFolderByServerRelativeUrl = function (folderRelativeUrl) {
	        return new folders_1.Folder(this, "getFolderByServerRelativeUrl('" + folderRelativeUrl + "')");
	    };
	    /**
	     * Get a file by server relative url
	     *
	     * @param fileRelativeUrl the server relative path to the file (including /sites/ if applicable)
	     */
	    Web.prototype.getFileByServerRelativeUrl = function (fileRelativeUrl) {
	        return new files_1.File(this, "getFileByServerRelativeUrl('" + fileRelativeUrl + "')");
	    };
	    /**
	     * Get a list by server relative url (list's root folder)
	     *
	     * @param listRelativeUrl the server relative path to the list's root folder (including /sites/ if applicable)
	     */
	    Web.prototype.getList = function (listRelativeUrl) {
	        return new lists_2.List(this, "getList('" + listRelativeUrl + "')");
	    };
	    /**
	     * Updates this web intance with the supplied properties
	     *
	     * @param properties A plain object hash of values to update for the web
	     */
	    Web.prototype.update = function (properties) {
	        var _this = this;
	        var postBody = JSON.stringify(util_1.Util.extend({
	            "__metadata": { "type": "SP.Web" },
	        }, properties));
	        return this.post({
	            body: postBody,
	            headers: {
	                "X-HTTP-Method": "MERGE",
	            },
	        }).then(function (data) {
	            return {
	                data: data,
	                web: _this,
	            };
	        });
	    };
	    /**
	     * Delete this web
	     *
	     */
	    Web.prototype.delete = function () {
	        return _super.prototype.delete.call(this);
	    };
	    /**
	     * Applies the theme specified by the contents of each of the files specified in the arguments to the site.
	     *
	     * @param colorPaletteUrl Server-relative URL of the color palette file.
	     * @param fontSchemeUrl Server-relative URL of the font scheme.
	     * @param backgroundImageUrl Server-relative URL of the background image.
	     * @param shareGenerated true to store the generated theme files in the root site, or false to store them in this site.
	     */
	    Web.prototype.applyTheme = function (colorPaletteUrl, fontSchemeUrl, backgroundImageUrl, shareGenerated) {
	        var postBody = JSON.stringify({
	            backgroundImageUrl: backgroundImageUrl,
	            colorPaletteUrl: colorPaletteUrl,
	            fontSchemeUrl: fontSchemeUrl,
	            shareGenerated: shareGenerated,
	        });
	        return this.clone(Web, "applytheme", true).post({ body: postBody });
	    };
	    /**
	     * Applies the specified site definition or site template to the Web site that has no template applied to it.
	     *
	     * @param template Name of the site definition or the name of the site template
	     */
	    Web.prototype.applyWebTemplate = function (template) {
	        var q = this.clone(Web, "applywebtemplate", true);
	        q.concat("(@t)");
	        q.query.add("@t", template);
	        return q.post();
	    };
	    /**
	     * Returns whether the current user has the given set of permissions.
	     *
	     * @param perms The high and low permission range.
	     */
	    Web.prototype.doesUserHavePermissions = function (perms) {
	        var q = this.clone(Web, "doesuserhavepermissions", true);
	        q.concat("(@p)");
	        q.query.add("@p", JSON.stringify(perms));
	        return q.get();
	    };
	    /**
	     * Checks whether the specified login name belongs to a valid user in the site. If the user doesn't exist, adds the user to the site.
	     *
	     * @param loginName The login name of the user (ex: i:0#.f|membership|user@domain.onmicrosoft.com)
	     */
	    Web.prototype.ensureUser = function (loginName) {
	        var postBody = JSON.stringify({
	            logonName: loginName,
	        });
	        return this.clone(Web, "ensureuser", true).post({ body: postBody }).then(function (data) {
	            return {
	                data: data,
	                user: new siteusers_1.SiteUser(odata_1.extractOdataId(data)),
	            };
	        });
	    };
	    /**
	     * Returns a collection of site templates available for the site.
	     *
	     * @param language The LCID of the site templates to get.
	     * @param true to include language-neutral site templates; otherwise false
	     */
	    Web.prototype.availableWebTemplates = function (language, includeCrossLanugage) {
	        if (language === void 0) { language = 1033; }
	        if (includeCrossLanugage === void 0) { includeCrossLanugage = true; }
	        return new queryable_1.QueryableCollection(this, "getavailablewebtemplates(lcid=" + language + ", doincludecrosslanguage=" + includeCrossLanugage + ")");
	    };
	    /**
	     * Returns the list gallery on the site.
	     *
	     * @param type The gallery type - WebTemplateCatalog = 111, WebPartCatalog = 113 ListTemplateCatalog = 114,
	     * MasterPageCatalog = 116, SolutionCatalog = 121, ThemeCatalog = 123, DesignCatalog = 124, AppDataCatalog = 125
	     */
	    Web.prototype.getCatalog = function (type) {
	        return this.clone(Web, "getcatalog(" + type + ")", true).select("Id").get().then(function (data) {
	            return new lists_2.List(odata_1.extractOdataId(data));
	        });
	    };
	    /**
	     * Returns the collection of changes from the change log that have occurred within the list, based on the specified query.
	     */
	    Web.prototype.getChanges = function (query) {
	        var postBody = JSON.stringify({ "query": util_1.Util.extend({ "__metadata": { "type": "SP.ChangeQuery" } }, query) });
	        return this.clone(Web, "getchanges", true).post({ body: postBody });
	    };
	    Object.defineProperty(Web.prototype, "customListTemplate", {
	        /**
	         * Gets the custom list templates for the site.
	         *
	         */
	        get: function () {
	            return new queryable_1.QueryableCollection(this, "getcustomlisttemplates");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    /**
	     * Returns the user corresponding to the specified member identifier for the current site.
	     *
	     * @param id The ID of the user.
	     */
	    Web.prototype.getUserById = function (id) {
	        return new siteusers_1.SiteUser(this, "getUserById(" + id + ")");
	    };
	    /**
	     * Returns the name of the image file for the icon that is used to represent the specified file.
	     *
	     * @param filename The file name. If this parameter is empty, the server returns an empty string.
	     * @param size The size of the icon: 16x16 pixels = 0, 32x32 pixels = 1.
	     * @param progId The ProgID of the application that was used to create the file, in the form OLEServerName.ObjectName
	     */
	    Web.prototype.mapToIcon = function (filename, size, progId) {
	        if (size === void 0) { size = 0; }
	        if (progId === void 0) { progId = ""; }
	        return this.clone(Web, "maptoicon(filename='" + filename + "', progid='" + progId + "', size=" + size + ")", true).get();
	    };
	    return Web;
	}(queryableshareable_1.QueryableShareableWeb));
	__decorate([
	    decorators_1.deprecated("This method will be removed in future releases. Please use the methods found in queryable securable.")
	], Web.prototype, "doesUserHavePermissions", null);
	exports.Web = Web;


/***/ },
/* 32 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	var __extends = (this && this.__extends) || (function () {
	    var extendStatics = Object.setPrototypeOf ||
	        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
	        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
	    return function (d, b) {
	        extendStatics(d, b);
	        function __() { this.constructor = d; }
	        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
	    };
	})();
	Object.defineProperty(exports, "__esModule", { value: true });
	var items_1 = __webpack_require__(33);
	var views_1 = __webpack_require__(45);
	var contenttypes_1 = __webpack_require__(43);
	var fields_1 = __webpack_require__(46);
	var forms_1 = __webpack_require__(47);
	var subscriptions_1 = __webpack_require__(48);
	var queryable_1 = __webpack_require__(22);
	var queryablesecurable_1 = __webpack_require__(35);
	var util_1 = __webpack_require__(12);
	var usercustomactions_1 = __webpack_require__(49);
	var odata_1 = __webpack_require__(23);
	var exceptions_1 = __webpack_require__(26);
	var folders_1 = __webpack_require__(40);
	/**
	 * Describes a collection of List objects
	 *
	 */
	var Lists = (function (_super) {
	    __extends(Lists, _super);
	    /**
	     * Creates a new instance of the Lists class
	     *
	     * @param baseUrl The url or Queryable which forms the parent of this fields collection
	     */
	    function Lists(baseUrl, path) {
	        if (path === void 0) { path = "lists"; }
	        return _super.call(this, baseUrl, path) || this;
	    }
	    /**
	     * Gets a list from the collection by title
	     *
	     * @param title The title of the list
	     */
	    Lists.prototype.getByTitle = function (title) {
	        return new List(this, "getByTitle('" + title + "')");
	    };
	    /**
	     * Gets a list from the collection by guid id
	     *
	     * @param id The Id of the list (GUID)
	     */
	    Lists.prototype.getById = function (id) {
	        var list = new List(this);
	        list.concat("('" + id + "')");
	        return list;
	    };
	    /**
	     * Adds a new list to the collection
	     *
	     * @param title The new list's title
	     * @param description The new list's description
	     * @param template The list template value
	     * @param enableContentTypes If true content types will be allowed and enabled, otherwise they will be disallowed and not enabled
	     * @param additionalSettings Will be passed as part of the list creation body
	     */
	    Lists.prototype.add = function (title, description, template, enableContentTypes, additionalSettings) {
	        var _this = this;
	        if (description === void 0) { description = ""; }
	        if (template === void 0) { template = 100; }
	        if (enableContentTypes === void 0) { enableContentTypes = false; }
	        if (additionalSettings === void 0) { additionalSettings = {}; }
	        var addSettings = util_1.Util.extend({
	            "AllowContentTypes": enableContentTypes,
	            "BaseTemplate": template,
	            "ContentTypesEnabled": enableContentTypes,
	            "Description": description,
	            "Title": title,
	            "__metadata": { "type": "SP.List" },
	        }, additionalSettings);
	        return this.post({ body: JSON.stringify(addSettings) }).then(function (data) {
	            return { data: data, list: _this.getByTitle(addSettings.Title) };
	        });
	    };
	    /**
	     * Ensures that the specified list exists in the collection (note: this method not supported for batching)
	     *
	     * @param title The new list's title
	     * @param description The new list's description
	     * @param template The list template value
	     * @param enableContentTypes If true content types will be allowed and enabled, otherwise they will be disallowed and not enabled
	     * @param additionalSettings Will be passed as part of the list creation body or used to update an existing list
	     */
	    Lists.prototype.ensure = function (title, description, template, enableContentTypes, additionalSettings) {
	        var _this = this;
	        if (description === void 0) { description = ""; }
	        if (template === void 0) { template = 100; }
	        if (enableContentTypes === void 0) { enableContentTypes = false; }
	        if (additionalSettings === void 0) { additionalSettings = {}; }
	        if (this.hasBatch) {
	            throw new exceptions_1.NotSupportedInBatchException("The ensure list method");
	        }
	        return new Promise(function (resolve, reject) {
	            var addOrUpdateSettings = util_1.Util.extend(additionalSettings, { Title: title, Description: description, ContentTypesEnabled: enableContentTypes }, true);
	            var list = _this.getByTitle(addOrUpdateSettings.Title);
	            list.get().then(function (_) {
	                list.update(addOrUpdateSettings).then(function (d) {
	                    resolve({ created: false, data: d, list: _this.getByTitle(addOrUpdateSettings.Title) });
	                }).catch(function (e) { return reject(e); });
	            }).catch(function (_) {
	                _this.add(title, description, template, enableContentTypes, addOrUpdateSettings).then(function (r) {
	                    resolve({ created: true, data: r.data, list: _this.getByTitle(addOrUpdateSettings.Title) });
	                }).catch(function (e) { return reject(e); });
	            });
	        });
	    };
	    /**
	     * Gets a list that is the default asset location for images or other files, which the users upload to their wiki pages.
	     */
	    Lists.prototype.ensureSiteAssetsLibrary = function () {
	        return this.clone(Lists, "ensuresiteassetslibrary", true).post().then(function (json) {
	            return new List(odata_1.extractOdataId(json));
	        });
	    };
	    /**
	     * Gets a list that is the default location for wiki pages.
	     */
	    Lists.prototype.ensureSitePagesLibrary = function () {
	        return this.clone(Lists, "ensuresitepageslibrary", true).post().then(function (json) {
	            return new List(odata_1.extractOdataId(json));
	        });
	    };
	    return Lists;
	}(queryable_1.QueryableCollection));
	exports.Lists = Lists;
	/**
	 * Describes a single List instance
	 *
	 */
	var List = (function (_super) {
	    __extends(List, _super);
	    function List() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    Object.defineProperty(List.prototype, "contentTypes", {
	        /**
	         * Gets the content types in this list
	         *
	         */
	        get: function () {
	            return new contenttypes_1.ContentTypes(this);
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(List.prototype, "items", {
	        /**
	         * Gets the items in this list
	         *
	         */
	        get: function () {
	            return new items_1.Items(this);
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(List.prototype, "views", {
	        /**
	         * Gets the views in this list
	         *
	         */
	        get: function () {
	            return new views_1.Views(this);
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(List.prototype, "fields", {
	        /**
	         * Gets the fields in this list
	         *
	         */
	        get: function () {
	            return new fields_1.Fields(this);
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(List.prototype, "forms", {
	        /**
	         * Gets the forms in this list
	         *
	         */
	        get: function () {
	            return new forms_1.Forms(this);
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(List.prototype, "defaultView", {
	        /**
	         * Gets the default view of this list
	         *
	         */
	        get: function () {
	            return new queryable_1.QueryableInstance(this, "DefaultView");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(List.prototype, "userCustomActions", {
	        /**
	         * Get all custom actions on a site collection
	         *
	         */
	        get: function () {
	            return new usercustomactions_1.UserCustomActions(this);
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(List.prototype, "effectiveBasePermissions", {
	        /**
	         * Gets the effective base permissions of this list
	         *
	         */
	        get: function () {
	            return new queryable_1.Queryable(this, "EffectiveBasePermissions");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(List.prototype, "eventReceivers", {
	        /**
	         * Gets the event receivers attached to this list
	         *
	         */
	        get: function () {
	            return new queryable_1.QueryableCollection(this, "EventReceivers");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(List.prototype, "relatedFields", {
	        /**
	         * Gets the related fields of this list
	         *
	         */
	        get: function () {
	            return new queryable_1.Queryable(this, "getRelatedFields");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(List.prototype, "informationRightsManagementSettings", {
	        /**
	         * Gets the IRM settings for this list
	         *
	         */
	        get: function () {
	            return new queryable_1.Queryable(this, "InformationRightsManagementSettings");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(List.prototype, "subscriptions", {
	        /**
	         * Gets the webhook subscriptions of this list
	         *
	         */
	        get: function () {
	            return new subscriptions_1.Subscriptions(this);
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(List.prototype, "rootFolder", {
	        /**
	         * The root folder of the list
	         */
	        get: function () {
	            return new folders_1.Folder(this, "rootFolder");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    /**
	     * Gets a view by view guid id
	     *
	     */
	    List.prototype.getView = function (viewId) {
	        return new views_1.View(this, "getView('" + viewId + "')");
	    };
	    /**
	     * Updates this list intance with the supplied properties
	     *
	     * @param properties A plain object hash of values to update for the list
	     * @param eTag Value used in the IF-Match header, by default "*"
	     */
	    /* tslint:disable no-string-literal */
	    List.prototype.update = function (properties, eTag) {
	        var _this = this;
	        if (eTag === void 0) { eTag = "*"; }
	        var postBody = JSON.stringify(util_1.Util.extend({
	            "__metadata": { "type": "SP.List" },
	        }, properties));
	        return this.post({
	            body: postBody,
	            headers: {
	                "IF-Match": eTag,
	                "X-HTTP-Method": "MERGE",
	            },
	        }).then(function (data) {
	            var retList = _this;
	            if (properties.hasOwnProperty("Title")) {
	                retList = _this.getParent(List, _this.parentUrl, "getByTitle('" + properties["Title"] + "')");
	            }
	            return {
	                data: data,
	                list: retList,
	            };
	        });
	    };
	    /* tslint:enable */
	    /**
	     * Delete this list
	     *
	     * @param eTag Value used in the IF-Match header, by default "*"
	     */
	    List.prototype.delete = function (eTag) {
	        if (eTag === void 0) { eTag = "*"; }
	        return this.post({
	            headers: {
	                "IF-Match": eTag,
	                "X-HTTP-Method": "DELETE",
	            },
	        });
	    };
	    /**
	     * Returns the collection of changes from the change log that have occurred within the list, based on the specified query.
	     */
	    List.prototype.getChanges = function (query) {
	        return this.clone(List, "getchanges", true).post({
	            body: JSON.stringify({ "query": util_1.Util.extend({ "__metadata": { "type": "SP.ChangeQuery" } }, query) }),
	        });
	    };
	    /**
	     * Returns a collection of items from the list based on the specified query.
	     *
	     * @param CamlQuery The Query schema of Collaborative Application Markup
	     * Language (CAML) is used in various ways within the context of Microsoft SharePoint Foundation
	     * to define queries against list data.
	     * see:
	     *
	     * https://msdn.microsoft.com/en-us/library/office/ms467521.aspx
	     *
	     * @param expands A URI with a $expand System Query Option indicates that Entries associated with
	     * the Entry or Collection of Entries identified by the Resource Path
	     * section of the URI must be represented inline (i.e. eagerly loaded).
	     * see:
	     *
	     * https://msdn.microsoft.com/en-us/library/office/fp142385.aspx
	     *
	     * http://www.odata.org/documentation/odata-version-2-0/uri-conventions/#ExpandSystemQueryOption
	     */
	    List.prototype.getItemsByCAMLQuery = function (query) {
	        var expands = [];
	        for (var _i = 1; _i < arguments.length; _i++) {
	            expands[_i - 1] = arguments[_i];
	        }
	        var q = this.clone(List, "getitems", true);
	        return q.expand.apply(q, expands).post({
	            body: JSON.stringify({ "query": util_1.Util.extend({ "__metadata": { "type": "SP.CamlQuery" } }, query) }),
	        });
	    };
	    /**
	     * See: https://msdn.microsoft.com/en-us/library/office/dn292554.aspx
	     */
	    List.prototype.getListItemChangesSinceToken = function (query) {
	        return this.clone(List, "getlistitemchangessincetoken", true).post({
	            body: JSON.stringify({ "query": util_1.Util.extend({ "__metadata": { "type": "SP.ChangeLogItemQuery" } }, query) }),
	        }, { parse: function (r) { return r.text(); } });
	    };
	    /**
	     * Moves the list to the Recycle Bin and returns the identifier of the new Recycle Bin item.
	     */
	    List.prototype.recycle = function () {
	        return this.clone(List, "recycle", true).post().then(function (data) {
	            if (data.hasOwnProperty("Recycle")) {
	                return data.Recycle;
	            }
	            else {
	                return data;
	            }
	        });
	    };
	    /**
	     * Renders list data based on the view xml provided
	     */
	    List.prototype.renderListData = function (viewXml) {
	        var q = this.clone(List, "renderlistdata(@viewXml)");
	        q.query.add("@viewXml", "'" + viewXml + "'");
	        return q.post().then(function (data) {
	            // data will be a string, so we parse it again
	            data = JSON.parse(data);
	            if (data.hasOwnProperty("RenderListData")) {
	                return data.RenderListData;
	            }
	            else {
	                return data;
	            }
	        });
	    };
	    /**
	     * Gets the field values and field schema attributes for a list item.
	     */
	    List.prototype.renderListFormData = function (itemId, formId, mode) {
	        return this.clone(List, "renderlistformdata(itemid=" + itemId + ", formid='" + formId + "', mode='" + mode + "')", true).post().then(function (data) {
	            // data will be a string, so we parse it again
	            data = JSON.parse(data);
	            if (data.hasOwnProperty("ListData")) {
	                return data.ListData;
	            }
	            else {
	                return data;
	            }
	        });
	    };
	    /**
	     * Reserves a list item ID for idempotent list item creation.
	     */
	    List.prototype.reserveListItemId = function () {
	        return this.clone(List, "reservelistitemid", true).post().then(function (data) {
	            if (data.hasOwnProperty("ReserveListItemId")) {
	                return data.ReserveListItemId;
	            }
	            else {
	                return data;
	            }
	        });
	    };
	    /**
	     * Returns the ListItemEntityTypeFullName for this list, used when adding/updating list items. Does not support batching.
	     *
	     */
	    List.prototype.getListItemEntityTypeFullName = function () {
	        return this.clone(List, null).select("ListItemEntityTypeFullName").getAs().then(function (o) { return o.ListItemEntityTypeFullName; });
	    };
	    return List;
	}(queryablesecurable_1.QueryableSecurable));
	exports.List = List;


/***/ },
/* 33 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	var __extends = (this && this.__extends) || (function () {
	    var extendStatics = Object.setPrototypeOf ||
	        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
	        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
	    return function (d, b) {
	        extendStatics(d, b);
	        function __() { this.constructor = d; }
	        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
	    };
	})();
	Object.defineProperty(exports, "__esModule", { value: true });
	var queryable_1 = __webpack_require__(22);
	var queryableshareable_1 = __webpack_require__(34);
	var folders_1 = __webpack_require__(40);
	var files_1 = __webpack_require__(41);
	var contenttypes_1 = __webpack_require__(43);
	var util_1 = __webpack_require__(12);
	var odata_1 = __webpack_require__(23);
	var attachmentfiles_1 = __webpack_require__(44);
	var lists_1 = __webpack_require__(32);
	/**
	 * Describes a collection of Item objects
	 *
	 */
	var Items = (function (_super) {
	    __extends(Items, _super);
	    /**
	     * Creates a new instance of the Items class
	     *
	     * @param baseUrl The url or Queryable which forms the parent of this fields collection
	     */
	    function Items(baseUrl, path) {
	        if (path === void 0) { path = "items"; }
	        return _super.call(this, baseUrl, path) || this;
	    }
	    /**
	     * Gets an Item by id
	     *
	     * @param id The integer id of the item to retrieve
	     */
	    Items.prototype.getById = function (id) {
	        var i = new Item(this);
	        i.concat("(" + id + ")");
	        return i;
	    };
	    /**
	     * Skips the specified number of items (https://msdn.microsoft.com/en-us/library/office/fp142385.aspx#sectionSection6)
	     *
	     * @param skip The starting id where the page should start, use with top to specify pages
	     */
	    Items.prototype.skip = function (skip) {
	        this._query.add("$skiptoken", encodeURIComponent("Paged=TRUE&p_ID=" + skip));
	        return this;
	    };
	    /**
	     * Gets a collection designed to aid in paging through data
	     *
	     */
	    Items.prototype.getPaged = function () {
	        return this.getAs(new PagedItemCollectionParser());
	    };
	    //
	    /**
	     * Adds a new item to the collection
	     *
	     * @param properties The new items's properties
	     */
	    Items.prototype.add = function (properties, listItemEntityTypeFullName) {
	        var _this = this;
	        if (properties === void 0) { properties = {}; }
	        if (listItemEntityTypeFullName === void 0) { listItemEntityTypeFullName = null; }
	        var removeDependency = this.addBatchDependency();
	        return this.ensureListItemEntityTypeName(listItemEntityTypeFullName).then(function (listItemEntityType) {
	            var postBody = JSON.stringify(util_1.Util.extend({
	                "__metadata": { "type": listItemEntityType },
	            }, properties));
	            var promise = _this.clone(Items, null, true).postAs({ body: postBody }).then(function (data) {
	                return {
	                    data: data,
	                    item: _this.getById(data.Id),
	                };
	            });
	            removeDependency();
	            return promise;
	        });
	    };
	    /**
	     * Ensures we have the proper list item entity type name, either from the value provided or from the list
	     *
	     * @param candidatelistItemEntityTypeFullName The potential type name
	     */
	    Items.prototype.ensureListItemEntityTypeName = function (candidatelistItemEntityTypeFullName) {
	        return candidatelistItemEntityTypeFullName ?
	            Promise.resolve(candidatelistItemEntityTypeFullName) :
	            this.getParent(lists_1.List).getListItemEntityTypeFullName();
	    };
	    return Items;
	}(queryable_1.QueryableCollection));
	exports.Items = Items;
	/**
	 * Descrines a single Item instance
	 *
	 */
	var Item = (function (_super) {
	    __extends(Item, _super);
	    function Item() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    Object.defineProperty(Item.prototype, "attachmentFiles", {
	        /**
	         * Gets the set of attachments for this item
	         *
	         */
	        get: function () {
	            return new attachmentfiles_1.AttachmentFiles(this);
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Item.prototype, "contentType", {
	        /**
	         * Gets the content type for this item
	         *
	         */
	        get: function () {
	            return new contenttypes_1.ContentType(this, "ContentType");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Item.prototype, "effectiveBasePermissions", {
	        /**
	         * Gets the effective base permissions for the item
	         *
	         */
	        get: function () {
	            return new queryable_1.Queryable(this, "EffectiveBasePermissions");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Item.prototype, "effectiveBasePermissionsForUI", {
	        /**
	         * Gets the effective base permissions for the item in a UI context
	         *
	         */
	        get: function () {
	            return new queryable_1.Queryable(this, "EffectiveBasePermissionsForUI");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Item.prototype, "fieldValuesAsHTML", {
	        /**
	         * Gets the field values for this list item in their HTML representation
	         *
	         */
	        get: function () {
	            return new queryable_1.QueryableInstance(this, "FieldValuesAsHTML");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Item.prototype, "fieldValuesAsText", {
	        /**
	         * Gets the field values for this list item in their text representation
	         *
	         */
	        get: function () {
	            return new queryable_1.QueryableInstance(this, "FieldValuesAsText");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Item.prototype, "fieldValuesForEdit", {
	        /**
	         * Gets the field values for this list item for use in editing controls
	         *
	         */
	        get: function () {
	            return new queryable_1.QueryableInstance(this, "FieldValuesForEdit");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Item.prototype, "folder", {
	        /**
	         * Gets the folder associated with this list item (if this item represents a folder)
	         *
	         */
	        get: function () {
	            return new folders_1.Folder(this, "folder");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Item.prototype, "file", {
	        /**
	         * Gets the folder associated with this list item (if this item represents a folder)
	         *
	         */
	        get: function () {
	            return new files_1.File(this, "file");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    /**
	     * Updates this list intance with the supplied properties
	     *
	     * @param properties A plain object hash of values to update for the list
	     * @param eTag Value used in the IF-Match header, by default "*"
	     */
	    Item.prototype.update = function (properties, eTag) {
	        var _this = this;
	        if (eTag === void 0) { eTag = "*"; }
	        return new Promise(function (resolve, reject) {
	            var removeDependency = _this.addBatchDependency();
	            var parentList = _this.getParent(queryable_1.QueryableInstance, _this.parentUrl.substr(0, _this.parentUrl.lastIndexOf("/")));
	            parentList.select("ListItemEntityTypeFullName").getAs().then(function (d) {
	                var postBody = JSON.stringify(util_1.Util.extend({
	                    "__metadata": { "type": d.ListItemEntityTypeFullName },
	                }, properties));
	                removeDependency();
	                return _this.post({
	                    body: postBody,
	                    headers: {
	                        "IF-Match": eTag,
	                        "X-HTTP-Method": "MERGE",
	                    },
	                }, new ItemUpdatedParser()).then(function (data) {
	                    resolve({
	                        data: data,
	                        item: _this,
	                    });
	                });
	            }).catch(function (e) { return reject(e); });
	        });
	    };
	    /**
	     * Delete this item
	     *
	     * @param eTag Value used in the IF-Match header, by default "*"
	     */
	    Item.prototype.delete = function (eTag) {
	        if (eTag === void 0) { eTag = "*"; }
	        return this.post({
	            headers: {
	                "IF-Match": eTag,
	                "X-HTTP-Method": "DELETE",
	            },
	        });
	    };
	    /**
	     * Moves the list item to the Recycle Bin and returns the identifier of the new Recycle Bin item.
	     */
	    Item.prototype.recycle = function () {
	        return this.clone(Item, "recycle", true).post();
	    };
	    /**
	     * Gets a string representation of the full URL to the WOPI frame.
	     * If there is no associated WOPI application, or no associated action, an empty string is returned.
	     *
	     * @param action Display mode: 0: view, 1: edit, 2: mobileView, 3: interactivePreview
	     */
	    Item.prototype.getWopiFrameUrl = function (action) {
	        if (action === void 0) { action = 0; }
	        var i = this.clone(Item, "getWOPIFrameUrl(@action)", true);
	        i._query.add("@action", action);
	        return i.post().then(function (data) {
	            return data.GetWOPIFrameUrl;
	        });
	    };
	    /**
	     * Validates and sets the values of the specified collection of fields for the list item.
	     *
	     * @param formValues The fields to change and their new values.
	     * @param newDocumentUpdate true if the list item is a document being updated after upload; otherwise false.
	     */
	    Item.prototype.validateUpdateListItem = function (formValues, newDocumentUpdate) {
	        if (newDocumentUpdate === void 0) { newDocumentUpdate = false; }
	        return this.clone(Item, "validateupdatelistitem", true).post({
	            body: JSON.stringify({ "formValues": formValues, bNewDocumentUpdate: newDocumentUpdate }),
	        });
	    };
	    return Item;
	}(queryableshareable_1.QueryableShareableItem));
	exports.Item = Item;
	/**
	 * Provides paging functionality for list items
	 */
	var PagedItemCollection = (function () {
	    function PagedItemCollection(nextUrl, results) {
	        this.nextUrl = nextUrl;
	        this.results = results;
	    }
	    Object.defineProperty(PagedItemCollection.prototype, "hasNext", {
	        /**
	         * If true there are more results available in the set, otherwise there are not
	         */
	        get: function () {
	            return typeof this.nextUrl === "string" && this.nextUrl.length > 0;
	        },
	        enumerable: true,
	        configurable: true
	    });
	    /**
	     * Gets the next set of results, or resolves to null if no results are available
	     */
	    PagedItemCollection.prototype.getNext = function () {
	        if (this.hasNext) {
	            var items = new Items(this.nextUrl, null);
	            return items.getPaged();
	        }
	        return new Promise(function (r) { return r(null); });
	    };
	    return PagedItemCollection;
	}());
	exports.PagedItemCollection = PagedItemCollection;
	var PagedItemCollectionParser = (function (_super) {
	    __extends(PagedItemCollectionParser, _super);
	    function PagedItemCollectionParser() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    PagedItemCollectionParser.prototype.parse = function (r) {
	        var _this = this;
	        return new Promise(function (resolve, reject) {
	            if (_this.handleError(r, reject)) {
	                r.json().then(function (json) {
	                    var nextUrl = json.hasOwnProperty("d") && json.d.hasOwnProperty("__next") ? json.d.__next : json["odata.nextLink"];
	                    resolve(new PagedItemCollection(nextUrl, _this.parseODataJSON(json)));
	                });
	            }
	        });
	    };
	    return PagedItemCollectionParser;
	}(odata_1.ODataParserBase));
	var ItemUpdatedParser = (function (_super) {
	    __extends(ItemUpdatedParser, _super);
	    function ItemUpdatedParser() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    ItemUpdatedParser.prototype.parse = function (r) {
	        var _this = this;
	        return new Promise(function (resolve, reject) {
	            if (_this.handleError(r, reject)) {
	                resolve({
	                    "odata.etag": r.headers.get("etag"),
	                });
	            }
	        });
	    };
	    return ItemUpdatedParser;
	}(odata_1.ODataParserBase));


/***/ },
/* 34 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	var __extends = (this && this.__extends) || (function () {
	    var extendStatics = Object.setPrototypeOf ||
	        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
	        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
	    return function (d, b) {
	        extendStatics(d, b);
	        function __() { this.constructor = d; }
	        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
	    };
	})();
	Object.defineProperty(exports, "__esModule", { value: true });
	var util_1 = __webpack_require__(12);
	var webs_1 = __webpack_require__(31);
	var odata_1 = __webpack_require__(23);
	var queryable_1 = __webpack_require__(22);
	var queryablesecurable_1 = __webpack_require__(35);
	var types_1 = __webpack_require__(39);
	/**
	 * Internal helper class used to augment classes to include sharing functionality
	 */
	var QueryableShareable = (function (_super) {
	    __extends(QueryableShareable, _super);
	    function QueryableShareable() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    /**
	     * Gets a sharing link for the supplied
	     *
	     * @param kind The kind of link to share
	     * @param expiration The optional expiration for this link
	     */
	    QueryableShareable.prototype.getShareLink = function (kind, expiration) {
	        if (expiration === void 0) { expiration = null; }
	        // date needs to be an ISO string or null
	        var expString = expiration !== null ? expiration.toISOString() : null;
	        // clone using the factory and send the request
	        return this.clone(QueryableShareable, "shareLink", true).postAs({
	            body: JSON.stringify({
	                request: {
	                    createLink: true,
	                    emailData: null,
	                    settings: {
	                        expiration: expString,
	                        linkKind: kind,
	                    },
	                },
	            }),
	        });
	    };
	    /**
	     * Shares this instance with the supplied users
	     *
	     * @param loginNames Resolved login names to share
	     * @param role The role
	     * @param requireSignin True to require the user is authenticated, otherwise false
	     * @param propagateAcl True to apply this share to all children
	     * @param emailData If supplied an email will be sent with the indicated properties
	     */
	    QueryableShareable.prototype.shareWith = function (loginNames, role, requireSignin, propagateAcl, emailData) {
	        var _this = this;
	        if (requireSignin === void 0) { requireSignin = true; }
	        if (propagateAcl === void 0) { propagateAcl = false; }
	        // handle the multiple input types
	        if (!Array.isArray(loginNames)) {
	            loginNames = [loginNames];
	        }
	        var userStr = JSON.stringify(loginNames.map(function (login) { return { Key: login }; }));
	        var roleFilter = role === types_1.SharingRole.Edit ? types_1.RoleType.Contributor : types_1.RoleType.Reader;
	        // start by looking up the role definition id we need to set the roleValue
	        return webs_1.Web.fromUrl(this.toUrl()).roleDefinitions.select("Id").filter("RoleTypeKind eq " + roleFilter).get().then(function (def) {
	            if (!Array.isArray(def) || def.length < 1) {
	                throw new Error("Could not locate a role defintion with RoleTypeKind " + roleFilter);
	            }
	            var postBody = {
	                includeAnonymousLinkInEmail: requireSignin,
	                peoplePickerInput: userStr,
	                propagateAcl: propagateAcl,
	                roleValue: "role:" + def[0].Id,
	                useSimplifiedRoles: true,
	            };
	            if (typeof emailData !== "undefined") {
	                postBody = util_1.Util.extend(postBody, {
	                    emailBody: emailData.body,
	                    emailSubject: typeof emailData.subject !== "undefined" ? "" : emailData.subject,
	                    sendEmail: true,
	                });
	            }
	            return _this.clone(QueryableShareable, "shareObject", true).postAs({
	                body: JSON.stringify(postBody),
	            });
	        });
	    };
	    /**
	     * Shares an object based on the supplied options
	     *
	     * @param options The set of options to send to the ShareObject method
	     * @param bypass If true any processing is skipped and the options are sent directly to the ShareObject method
	     */
	    QueryableShareable.prototype.shareObject = function (options, bypass) {
	        var _this = this;
	        if (bypass === void 0) { bypass = false; }
	        if (bypass) {
	            // if the bypass flag is set send the supplied parameters directly to the service
	            return this.sendShareObjectRequest(options);
	        }
	        // extend our options with some defaults
	        options = util_1.Util.extend(options, {
	            group: null,
	            includeAnonymousLinkInEmail: false,
	            propagateAcl: false,
	            useSimplifiedRoles: true,
	        }, true);
	        return this.getRoleValue(options.role, options.group).then(function (roleValue) {
	            // handle the multiple input types
	            if (!Array.isArray(options.loginNames)) {
	                options.loginNames = [options.loginNames];
	            }
	            var userStr = JSON.stringify(options.loginNames.map(function (login) { return { Key: login }; }));
	            var postBody = {
	                peoplePickerInput: userStr,
	                roleValue: roleValue,
	                url: options.url,
	            };
	            if (typeof options.emailData !== "undefined" && options.emailData !== null) {
	                postBody = util_1.Util.extend(postBody, {
	                    emailBody: options.emailData.body,
	                    emailSubject: typeof options.emailData.subject !== "undefined" ? "Shared for you." : options.emailData.subject,
	                    sendEmail: true,
	                });
	            }
	            return _this.sendShareObjectRequest(postBody);
	        });
	    };
	    /**
	     * Calls the web's UnshareObject method
	     *
	     * @param url The url of the object to unshare
	     */
	    QueryableShareable.prototype.unshareObjectWeb = function (url) {
	        return this.clone(QueryableShareable, "unshareObject", true).postAs({
	            body: JSON.stringify({
	                url: url,
	            }),
	        });
	    };
	    /**
	     * Checks Permissions on the list of Users and returns back role the users have on the Item.
	     *
	     * @param recipients The array of Entities for which Permissions need to be checked.
	     */
	    QueryableShareable.prototype.checkPermissions = function (recipients) {
	        return this.clone(QueryableShareable, "checkPermissions", true).postAs({
	            body: JSON.stringify({
	                recipients: recipients,
	            }),
	        });
	    };
	    /**
	     * Get Sharing Information.
	     *
	     * @param request The SharingInformationRequest Object.
	     */
	    QueryableShareable.prototype.getSharingInformation = function (request) {
	        if (request === void 0) { request = null; }
	        return this.clone(QueryableShareable, "getSharingInformation", true).postAs({
	            body: JSON.stringify({
	                request: request,
	            }),
	        });
	    };
	    /**
	     * Gets the sharing settings of an item.
	     *
	     * @param useSimplifiedRoles Determines whether to use simplified roles.
	     */
	    QueryableShareable.prototype.getObjectSharingSettings = function (useSimplifiedRoles) {
	        if (useSimplifiedRoles === void 0) { useSimplifiedRoles = true; }
	        return this.clone(QueryableShareable, "getObjectSharingSettings", true).postAs({
	            body: JSON.stringify({
	                useSimplifiedRoles: useSimplifiedRoles,
	            }),
	        });
	    };
	    /**
	     * Unshares this object
	     */
	    QueryableShareable.prototype.unshareObject = function () {
	        return this.clone(QueryableShareable, "unshareObject", true).postAs();
	    };
	    /**
	     * Deletes a link by type
	     *
	     * @param kind Deletes a sharing link by the kind of link
	     */
	    QueryableShareable.prototype.deleteLinkByKind = function (kind) {
	        return this.clone(QueryableShareable, "deleteLinkByKind", true).post({
	            body: JSON.stringify({ linkKind: kind }),
	        });
	    };
	    /**
	     * Removes the specified link to the item.
	     *
	     * @param kind The kind of link to be deleted.
	     * @param shareId
	     */
	    QueryableShareable.prototype.unshareLink = function (kind, shareId) {
	        if (shareId === void 0) { shareId = "00000000-0000-0000-0000-000000000000"; }
	        return this.clone(QueryableShareable, "unshareLink", true).post({
	            body: JSON.stringify({ linkKind: kind, shareId: shareId }),
	        });
	    };
	    /**
	     * Calculates the roleValue string used in the sharing query
	     *
	     * @param role The Sharing Role
	     * @param group The Group type
	     */
	    QueryableShareable.prototype.getRoleValue = function (role, group) {
	        // we will give group precedence, because we had to make a choice
	        if (typeof group !== "undefined" && group !== null) {
	            switch (group) {
	                case types_1.RoleType.Contributor:
	                    return webs_1.Web.fromUrl(this.toUrl()).associatedMemberGroup.select("Id").getAs().then(function (g) { return "group: " + g.Id; });
	                case types_1.RoleType.Reader:
	                case types_1.RoleType.Guest:
	                    return webs_1.Web.fromUrl(this.toUrl()).associatedVisitorGroup.select("Id").getAs().then(function (g) { return "group: " + g.Id; });
	                default:
	                    throw new Error("Could not determine role value for supplied value. Contributor, Reader, and Guest are supported");
	            }
	        }
	        else {
	            var roleFilter = role === types_1.SharingRole.Edit ? types_1.RoleType.Contributor : types_1.RoleType.Reader;
	            return webs_1.Web.fromUrl(this.toUrl()).roleDefinitions.select("Id").top(1).filter("RoleTypeKind eq " + roleFilter).getAs().then(function (def) {
	                if (def.length < 1) {
	                    throw new Error("Could not locate associated role definition for supplied role. Edit and View are supported");
	                }
	                return "role: " + def[0].Id;
	            });
	        }
	    };
	    QueryableShareable.prototype.getShareObjectWeb = function (candidate) {
	        return Promise.resolve(webs_1.Web.fromUrl(candidate, "/_api/SP.Web.ShareObject"));
	    };
	    QueryableShareable.prototype.sendShareObjectRequest = function (options) {
	        return this.getShareObjectWeb(this.toUrl()).then(function (web) {
	            return web.expand("UsersWithAccessRequests", "GroupsSharedWith").as(QueryableShareable).post({
	                body: JSON.stringify(options),
	            });
	        });
	    };
	    return QueryableShareable;
	}(queryable_1.Queryable));
	exports.QueryableShareable = QueryableShareable;
	var QueryableShareableWeb = (function (_super) {
	    __extends(QueryableShareableWeb, _super);
	    function QueryableShareableWeb() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    /**
	     * Shares this web with the supplied users
	     * @param loginNames The resolved login names to share
	     * @param role The role to share this web
	     * @param emailData Optional email data
	     */
	    QueryableShareableWeb.prototype.shareWith = function (loginNames, role, emailData) {
	        var _this = this;
	        if (role === void 0) { role = types_1.SharingRole.View; }
	        var dependency = this.addBatchDependency();
	        return webs_1.Web.fromUrl(this.toUrl(), "/_api/web/url").get().then(function (url) {
	            dependency();
	            return _this.shareObject(util_1.Util.combinePaths(url, "/_layouts/15/aclinv.aspx?forSharing=1&mbypass=1"), loginNames, role, emailData);
	        });
	    };
	    /**
	     * Provides direct access to the static web.ShareObject method
	     *
	     * @param url The url to share
	     * @param loginNames Resolved loginnames string[] of a single login name string
	     * @param roleValue Role value
	     * @param emailData Optional email data
	     * @param groupId Optional group id
	     * @param propagateAcl
	     * @param includeAnonymousLinkInEmail
	     * @param useSimplifiedRoles
	     */
	    QueryableShareableWeb.prototype.shareObject = function (url, loginNames, role, emailData, group, propagateAcl, includeAnonymousLinkInEmail, useSimplifiedRoles) {
	        if (propagateAcl === void 0) { propagateAcl = false; }
	        if (includeAnonymousLinkInEmail === void 0) { includeAnonymousLinkInEmail = false; }
	        if (useSimplifiedRoles === void 0) { useSimplifiedRoles = true; }
	        return this.clone(QueryableShareable, null, true).shareObject({
	            emailData: emailData,
	            group: group,
	            includeAnonymousLinkInEmail: includeAnonymousLinkInEmail,
	            loginNames: loginNames,
	            propagateAcl: propagateAcl,
	            role: role,
	            url: url,
	            useSimplifiedRoles: useSimplifiedRoles,
	        });
	    };
	    /**
	     * Supplies a method to pass any set of arguments to ShareObject
	     *
	     * @param options The set of options to send to ShareObject
	     */
	    QueryableShareableWeb.prototype.shareObjectRaw = function (options) {
	        return this.clone(QueryableShareable, null, true).shareObject(options, true);
	    };
	    /**
	     * Unshares the object
	     *
	     * @param url The url of the object to stop sharing
	     */
	    QueryableShareableWeb.prototype.unshareObject = function (url) {
	        return this.clone(QueryableShareable, null, true).unshareObjectWeb(url);
	    };
	    return QueryableShareableWeb;
	}(queryablesecurable_1.QueryableSecurable));
	exports.QueryableShareableWeb = QueryableShareableWeb;
	var QueryableShareableItem = (function (_super) {
	    __extends(QueryableShareableItem, _super);
	    function QueryableShareableItem() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    /**
	     * Gets a link suitable for sharing for this item
	     *
	     * @param kind The type of link to share
	     * @param expiration The optional expiration date
	     */
	    QueryableShareableItem.prototype.getShareLink = function (kind, expiration) {
	        if (kind === void 0) { kind = types_1.SharingLinkKind.OrganizationView; }
	        if (expiration === void 0) { expiration = null; }
	        return this.clone(QueryableShareable, null, true).getShareLink(kind, expiration);
	    };
	    /**
	     * Shares this item with one or more users
	     *
	     * @param loginNames string or string[] of resolved login names to which this item will be shared
	     * @param role The role (View | Edit) applied to the share
	     * @param emailData Optional, if inlucded an email will be sent. Note subject currently has no effect.
	     */
	    QueryableShareableItem.prototype.shareWith = function (loginNames, role, requireSignin, emailData) {
	        if (role === void 0) { role = types_1.SharingRole.View; }
	        if (requireSignin === void 0) { requireSignin = true; }
	        return this.clone(QueryableShareable, null, true).shareWith(loginNames, role, requireSignin, false, emailData);
	    };
	    /**
	     * Checks Permissions on the list of Users and returns back role the users have on the Item.
	     *
	     * @param recipients The array of Entities for which Permissions need to be checked.
	     */
	    QueryableShareableItem.prototype.checkSharingPermissions = function (recipients) {
	        return this.clone(QueryableShareable, null, true).checkPermissions(recipients);
	    };
	    /**
	     * Get Sharing Information.
	     *
	     * @param request The SharingInformationRequest Object.
	     */
	    QueryableShareableItem.prototype.getSharingInformation = function (request) {
	        if (request === void 0) { request = null; }
	        return this.clone(QueryableShareable, null, true).getSharingInformation(request);
	    };
	    /**
	     * Gets the sharing settings of an item.
	     *
	     * @param useSimplifiedRoles Determines whether to use simplified roles.
	     */
	    QueryableShareableItem.prototype.getObjectSharingSettings = function (useSimplifiedRoles) {
	        if (useSimplifiedRoles === void 0) { useSimplifiedRoles = true; }
	        return this.clone(QueryableShareable, null, true).getObjectSharingSettings(useSimplifiedRoles);
	    };
	    /**
	     * Unshare this item
	     */
	    QueryableShareableItem.prototype.unshare = function () {
	        return this.clone(QueryableShareable, null, true).unshareObject();
	    };
	    /**
	     * Deletes a sharing link by kind
	     *
	     * @param kind Deletes a sharing link by the kind of link
	     */
	    QueryableShareableItem.prototype.deleteSharingLinkByKind = function (kind) {
	        return this.clone(QueryableShareable, null, true).deleteLinkByKind(kind);
	    };
	    /**
	     * Removes the specified link to the item.
	     *
	     * @param kind The kind of link to be deleted.
	     * @param shareId
	     */
	    QueryableShareableItem.prototype.unshareLink = function (kind, shareId) {
	        return this.clone(QueryableShareable, null, true).unshareLink(kind, shareId);
	    };
	    return QueryableShareableItem;
	}(queryablesecurable_1.QueryableSecurable));
	exports.QueryableShareableItem = QueryableShareableItem;
	var FileFolderShared = (function (_super) {
	    __extends(FileFolderShared, _super);
	    function FileFolderShared() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    /**
	     * Gets a link suitable for sharing
	     *
	     * @param kind The kind of link to get
	     * @param expiration Optional, an expiration for this link
	     */
	    FileFolderShared.prototype.getShareLink = function (kind, expiration) {
	        if (kind === void 0) { kind = types_1.SharingLinkKind.OrganizationView; }
	        if (expiration === void 0) { expiration = null; }
	        var dependency = this.addBatchDependency();
	        return this.getShareable().then(function (shareable) {
	            dependency();
	            return shareable.getShareLink(kind, expiration);
	        });
	    };
	    /**
	         * Checks Permissions on the list of Users and returns back role the users have on the Item.
	         *
	         * @param recipients The array of Entities for which Permissions need to be checked.
	         */
	    FileFolderShared.prototype.checkSharingPermissions = function (recipients) {
	        var dependency = this.addBatchDependency();
	        return this.getShareable().then(function (shareable) {
	            dependency();
	            return shareable.checkPermissions(recipients);
	        });
	    };
	    /**
	     * Get Sharing Information.
	     *
	     * @param request The SharingInformationRequest Object.
	     */
	    FileFolderShared.prototype.getSharingInformation = function (request) {
	        if (request === void 0) { request = null; }
	        var dependency = this.addBatchDependency();
	        return this.getShareable().then(function (shareable) {
	            dependency();
	            return shareable.getSharingInformation(request);
	        });
	    };
	    /**
	     * Gets the sharing settings of an item.
	     *
	     * @param useSimplifiedRoles Determines whether to use simplified roles.
	     */
	    FileFolderShared.prototype.getObjectSharingSettings = function (useSimplifiedRoles) {
	        if (useSimplifiedRoles === void 0) { useSimplifiedRoles = true; }
	        var dependency = this.addBatchDependency();
	        return this.getShareable().then(function (shareable) {
	            dependency();
	            return shareable.getObjectSharingSettings(useSimplifiedRoles);
	        });
	    };
	    /**
	     * Unshare this item
	     */
	    FileFolderShared.prototype.unshare = function () {
	        var dependency = this.addBatchDependency();
	        return this.getShareable().then(function (shareable) {
	            dependency();
	            return shareable.unshareObject();
	        });
	    };
	    /**
	     * Deletes a sharing link by the kind of link
	     *
	     * @param kind The kind of link to be deleted.
	     */
	    FileFolderShared.prototype.deleteSharingLinkByKind = function (kind) {
	        var dependency = this.addBatchDependency();
	        return this.getShareable().then(function (shareable) {
	            dependency();
	            return shareable.deleteLinkByKind(kind);
	        });
	    };
	    /**
	     * Removes the specified link to the item.
	     *
	     * @param kind The kind of link to be deleted.
	     * @param shareId The share id to delete
	     */
	    FileFolderShared.prototype.unshareLink = function (kind, shareId) {
	        var dependency = this.addBatchDependency();
	        return this.getShareable().then(function (shareable) {
	            dependency();
	            return shareable.unshareLink(kind, shareId);
	        });
	    };
	    /**
	     * For files and folders we need to use the associated item end point
	     */
	    FileFolderShared.prototype.getShareable = function () {
	        var _this = this;
	        // sharing only works on the item end point, not the file one - so we create a folder instance with the item url internally
	        return this.clone(QueryableShareableFile, "listItemAllFields", false).select("odata.editlink").get().then(function (d) {
	            var shareable = new QueryableShareable(odata_1.getEntityUrl(d));
	            // we need to handle batching
	            if (_this.hasBatch) {
	                shareable = shareable.inBatch(_this.batch);
	            }
	            return shareable;
	        });
	    };
	    return FileFolderShared;
	}(queryable_1.QueryableInstance));
	exports.FileFolderShared = FileFolderShared;
	var QueryableShareableFile = (function (_super) {
	    __extends(QueryableShareableFile, _super);
	    function QueryableShareableFile() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    /**
	     * Shares this item with one or more users
	     *
	     * @param loginNames string or string[] of resolved login names to which this item will be shared
	     * @param role The role (View | Edit) applied to the share
	     * @param shareEverything Share everything in this folder, even items with unique permissions.
	     * @param requireSignin If true the user must signin to view link, otherwise anyone with the link can access the resource
	     * @param emailData Optional, if inlucded an email will be sent. Note subject currently has no effect.
	     */
	    QueryableShareableFile.prototype.shareWith = function (loginNames, role, requireSignin, emailData) {
	        if (role === void 0) { role = types_1.SharingRole.View; }
	        if (requireSignin === void 0) { requireSignin = true; }
	        var dependency = this.addBatchDependency();
	        return this.getShareable().then(function (shareable) {
	            dependency();
	            return shareable.shareWith(loginNames, role, requireSignin, false, emailData);
	        });
	    };
	    return QueryableShareableFile;
	}(FileFolderShared));
	exports.QueryableShareableFile = QueryableShareableFile;
	var QueryableShareableFolder = (function (_super) {
	    __extends(QueryableShareableFolder, _super);
	    function QueryableShareableFolder() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    /**
	     * Shares this item with one or more users
	     *
	     * @param loginNames string or string[] of resolved login names to which this item will be shared
	     * @param role The role (View | Edit) applied to the share
	     * @param shareEverything Share everything in this folder, even items with unique permissions.
	     * @param requireSignin If true the user must signin to view link, otherwise anyone with the link can access the resource
	     * @param emailData Optional, if inlucded an email will be sent. Note subject currently has no effect.
	     */
	    QueryableShareableFolder.prototype.shareWith = function (loginNames, role, requireSignin, shareEverything, emailData) {
	        if (role === void 0) { role = types_1.SharingRole.View; }
	        if (requireSignin === void 0) { requireSignin = true; }
	        if (shareEverything === void 0) { shareEverything = false; }
	        var dependency = this.addBatchDependency();
	        return this.getShareable().then(function (shareable) {
	            dependency();
	            return shareable.shareWith(loginNames, role, requireSignin, shareEverything, emailData);
	        });
	    };
	    return QueryableShareableFolder;
	}(FileFolderShared));
	exports.QueryableShareableFolder = QueryableShareableFolder;


/***/ },
/* 35 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	var __extends = (this && this.__extends) || (function () {
	    var extendStatics = Object.setPrototypeOf ||
	        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
	        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
	    return function (d, b) {
	        extendStatics(d, b);
	        function __() { this.constructor = d; }
	        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
	    };
	})();
	Object.defineProperty(exports, "__esModule", { value: true });
	var webs_1 = __webpack_require__(31);
	var roles_1 = __webpack_require__(36);
	var types_1 = __webpack_require__(39);
	var queryable_1 = __webpack_require__(22);
	var QueryableSecurable = (function (_super) {
	    __extends(QueryableSecurable, _super);
	    function QueryableSecurable() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    Object.defineProperty(QueryableSecurable.prototype, "roleAssignments", {
	        /**
	         * Gets the set of role assignments for this item
	         *
	         */
	        get: function () {
	            return new roles_1.RoleAssignments(this);
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(QueryableSecurable.prototype, "firstUniqueAncestorSecurableObject", {
	        /**
	         * Gets the closest securable up the security hierarchy whose permissions are applied to this list item
	         *
	         */
	        get: function () {
	            return new queryable_1.QueryableInstance(this, "FirstUniqueAncestorSecurableObject");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    /**
	     * Gets the effective permissions for the user supplied
	     *
	     * @param loginName The claims username for the user (ex: i:0#.f|membership|user@domain.com)
	     */
	    QueryableSecurable.prototype.getUserEffectivePermissions = function (loginName) {
	        var q = this.clone(queryable_1.Queryable, "getUserEffectivePermissions(@user)", true);
	        q.query.add("@user", "'" + encodeURIComponent(loginName) + "'");
	        return q.getAs();
	    };
	    /**
	     * Gets the effective permissions for the current user
	     */
	    QueryableSecurable.prototype.getCurrentUserEffectivePermissions = function () {
	        var _this = this;
	        var w = webs_1.Web.fromUrl(this.toUrl());
	        return w.currentUser.select("LoginName").getAs().then(function (user) {
	            return _this.getUserEffectivePermissions(user.LoginName);
	        });
	    };
	    /**
	     * Breaks the security inheritance at this level optinally copying permissions and clearing subscopes
	     *
	     * @param copyRoleAssignments If true the permissions are copied from the current parent scope
	     * @param clearSubscopes Optional. true to make all child securable objects inherit role assignments from the current object
	     */
	    QueryableSecurable.prototype.breakRoleInheritance = function (copyRoleAssignments, clearSubscopes) {
	        if (copyRoleAssignments === void 0) { copyRoleAssignments = false; }
	        if (clearSubscopes === void 0) { clearSubscopes = false; }
	        return this.clone(QueryableSecurable, "breakroleinheritance(copyroleassignments=" + copyRoleAssignments + ", clearsubscopes=" + clearSubscopes + ")", true).post();
	    };
	    /**
	     * Removes the local role assignments so that it re-inherit role assignments from the parent object.
	     *
	     */
	    QueryableSecurable.prototype.resetRoleInheritance = function () {
	        return this.clone(QueryableSecurable, "resetroleinheritance", true).post();
	    };
	    /**
	     * Determines if a given user has the appropriate permissions
	     *
	     * @param loginName The user to check
	     * @param permission The permission being checked
	     */
	    QueryableSecurable.prototype.userHasPermissions = function (loginName, permission) {
	        var _this = this;
	        return this.getUserEffectivePermissions(loginName).then(function (perms) {
	            return _this.hasPermissions(perms, permission);
	        });
	    };
	    /**
	     * Determines if the current user has the requested permissions
	     *
	     * @param permission The permission we wish to check
	     */
	    QueryableSecurable.prototype.currentUserHasPermissions = function (permission) {
	        var _this = this;
	        return this.getCurrentUserEffectivePermissions().then(function (perms) {
	            return _this.hasPermissions(perms, permission);
	        });
	    };
	    /**
	     * Taken from sp.js, checks the supplied permissions against the mask
	     *
	     * @param value The security principal's permissions on the given object
	     * @param perm The permission checked against the value
	     */
	    /* tslint:disable:no-bitwise */
	    QueryableSecurable.prototype.hasPermissions = function (value, perm) {
	        if (!perm) {
	            return true;
	        }
	        if (perm === types_1.PermissionKind.FullMask) {
	            return (value.High & 32767) === 32767 && value.Low === 65535;
	        }
	        perm = perm - 1;
	        var num = 1;
	        if (perm >= 0 && perm < 32) {
	            num = num << perm;
	            return 0 !== (value.Low & num);
	        }
	        else if (perm >= 32 && perm < 64) {
	            num = num << perm - 32;
	            return 0 !== (value.High & num);
	        }
	        return false;
	    };
	    return QueryableSecurable;
	}(queryable_1.QueryableInstance));
	exports.QueryableSecurable = QueryableSecurable;


/***/ },
/* 36 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	var __extends = (this && this.__extends) || (function () {
	    var extendStatics = Object.setPrototypeOf ||
	        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
	        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
	    return function (d, b) {
	        extendStatics(d, b);
	        function __() { this.constructor = d; }
	        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
	    };
	})();
	Object.defineProperty(exports, "__esModule", { value: true });
	var queryable_1 = __webpack_require__(22);
	var sitegroups_1 = __webpack_require__(37);
	var util_1 = __webpack_require__(12);
	/**
	 * Describes a set of role assignments for the current scope
	 *
	 */
	var RoleAssignments = (function (_super) {
	    __extends(RoleAssignments, _super);
	    /**
	     * Creates a new instance of the RoleAssignments class
	     *
	     * @param baseUrl The url or Queryable which forms the parent of this fields collection
	     */
	    function RoleAssignments(baseUrl, path) {
	        if (path === void 0) { path = "roleassignments"; }
	        return _super.call(this, baseUrl, path) || this;
	    }
	    /**
	     * Adds a new role assignment with the specified principal and role definitions to the collection.
	     *
	     * @param principalId The ID of the user or group to assign permissions to
	     * @param roleDefId The ID of the role definition that defines the permissions to assign
	     *
	     */
	    RoleAssignments.prototype.add = function (principalId, roleDefId) {
	        return this.clone(RoleAssignments, "addroleassignment(principalid=" + principalId + ", roledefid=" + roleDefId + ")", true).post();
	    };
	    /**
	     * Removes the role assignment with the specified principal and role definition from the collection
	     *
	     * @param principalId The ID of the user or group in the role assignment.
	     * @param roleDefId The ID of the role definition in the role assignment
	     *
	     */
	    RoleAssignments.prototype.remove = function (principalId, roleDefId) {
	        return this.clone(RoleAssignments, "removeroleassignment(principalid=" + principalId + ", roledefid=" + roleDefId + ")", true).post();
	    };
	    /**
	     * Gets the role assignment associated with the specified principal ID from the collection.
	     *
	     * @param id The id of the role assignment
	     */
	    RoleAssignments.prototype.getById = function (id) {
	        var ra = new RoleAssignment(this);
	        ra.concat("(" + id + ")");
	        return ra;
	    };
	    return RoleAssignments;
	}(queryable_1.QueryableCollection));
	exports.RoleAssignments = RoleAssignments;
	var RoleAssignment = (function (_super) {
	    __extends(RoleAssignment, _super);
	    function RoleAssignment() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    Object.defineProperty(RoleAssignment.prototype, "groups", {
	        get: function () {
	            return new sitegroups_1.SiteGroups(this, "groups");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(RoleAssignment.prototype, "bindings", {
	        /**
	         * Get the role definition bindings for this role assignment
	         *
	         */
	        get: function () {
	            return new RoleDefinitionBindings(this);
	        },
	        enumerable: true,
	        configurable: true
	    });
	    /**
	     * Delete this role assignment
	     *
	     */
	    RoleAssignment.prototype.delete = function () {
	        return this.post({
	            headers: {
	                "X-HTTP-Method": "DELETE",
	            },
	        });
	    };
	    return RoleAssignment;
	}(queryable_1.QueryableInstance));
	exports.RoleAssignment = RoleAssignment;
	var RoleDefinitions = (function (_super) {
	    __extends(RoleDefinitions, _super);
	    /**
	     * Creates a new instance of the RoleDefinitions class
	     *
	     * @param baseUrl The url or Queryable which forms the parent of this fields collection
	     * @param path
	     *
	     */
	    function RoleDefinitions(baseUrl, path) {
	        if (path === void 0) { path = "roledefinitions"; }
	        return _super.call(this, baseUrl, path) || this;
	    }
	    /**
	     * Gets the role definition with the specified ID from the collection.
	     *
	     * @param id The ID of the role definition.
	     *
	     */
	    RoleDefinitions.prototype.getById = function (id) {
	        return new RoleDefinition(this, "getById(" + id + ")");
	    };
	    /**
	     * Gets the role definition with the specified name.
	     *
	     * @param name The name of the role definition.
	     *
	     */
	    RoleDefinitions.prototype.getByName = function (name) {
	        return new RoleDefinition(this, "getbyname('" + name + "')");
	    };
	    /**
	     * Gets the role definition with the specified type.
	     *
	     * @param name The name of the role definition.
	     *
	     */
	    RoleDefinitions.prototype.getByType = function (roleTypeKind) {
	        return new RoleDefinition(this, "getbytype(" + roleTypeKind + ")");
	    };
	    /**
	     * Create a role definition
	     *
	     * @param name The new role definition's name
	     * @param description The new role definition's description
	     * @param order The order in which the role definition appears
	     * @param basePermissions The permissions mask for this role definition
	     *
	     */
	    RoleDefinitions.prototype.add = function (name, description, order, basePermissions) {
	        var _this = this;
	        var postBody = JSON.stringify({
	            BasePermissions: util_1.Util.extend({ __metadata: { type: "SP.BasePermissions" } }, basePermissions),
	            Description: description,
	            Name: name,
	            Order: order,
	            __metadata: { "type": "SP.RoleDefinition" },
	        });
	        return this.post({ body: postBody }).then(function (data) {
	            return {
	                data: data,
	                definition: _this.getById(data.Id),
	            };
	        });
	    };
	    return RoleDefinitions;
	}(queryable_1.QueryableCollection));
	exports.RoleDefinitions = RoleDefinitions;
	var RoleDefinition = (function (_super) {
	    __extends(RoleDefinition, _super);
	    function RoleDefinition() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    /**
	     * Updates this web intance with the supplied properties
	     *
	     * @param properties A plain object hash of values to update for the web
	     */
	    /* tslint:disable no-string-literal */
	    RoleDefinition.prototype.update = function (properties) {
	        var _this = this;
	        if (typeof properties.hasOwnProperty("BasePermissions") !== "undefined") {
	            properties["BasePermissions"] = util_1.Util.extend({ __metadata: { type: "SP.BasePermissions" } }, properties["BasePermissions"]);
	        }
	        var postBody = JSON.stringify(util_1.Util.extend({
	            "__metadata": { "type": "SP.RoleDefinition" },
	        }, properties));
	        return this.post({
	            body: postBody,
	            headers: {
	                "X-HTTP-Method": "MERGE",
	            },
	        }).then(function (data) {
	            var retDef = _this;
	            if (properties.hasOwnProperty("Name")) {
	                var parent_1 = _this.getParent(RoleDefinitions, _this.parentUrl, "");
	                retDef = parent_1.getByName(properties["Name"]);
	            }
	            return {
	                data: data,
	                definition: retDef,
	            };
	        });
	    };
	    /* tslint:enable */
	    /**
	     * Delete this role definition
	     *
	     */
	    RoleDefinition.prototype.delete = function () {
	        return this.post({
	            headers: {
	                "X-HTTP-Method": "DELETE",
	            },
	        });
	    };
	    return RoleDefinition;
	}(queryable_1.QueryableInstance));
	exports.RoleDefinition = RoleDefinition;
	var RoleDefinitionBindings = (function (_super) {
	    __extends(RoleDefinitionBindings, _super);
	    function RoleDefinitionBindings(baseUrl, path) {
	        if (path === void 0) { path = "roledefinitionbindings"; }
	        return _super.call(this, baseUrl, path) || this;
	    }
	    return RoleDefinitionBindings;
	}(queryable_1.QueryableCollection));
	exports.RoleDefinitionBindings = RoleDefinitionBindings;


/***/ },
/* 37 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	var __extends = (this && this.__extends) || (function () {
	    var extendStatics = Object.setPrototypeOf ||
	        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
	        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
	    return function (d, b) {
	        extendStatics(d, b);
	        function __() { this.constructor = d; }
	        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
	    };
	})();
	Object.defineProperty(exports, "__esModule", { value: true });
	var queryable_1 = __webpack_require__(22);
	var siteusers_1 = __webpack_require__(38);
	var util_1 = __webpack_require__(12);
	/**
	 * Principal Type enum
	 *
	 */
	var PrincipalType;
	(function (PrincipalType) {
	    PrincipalType[PrincipalType["None"] = 0] = "None";
	    PrincipalType[PrincipalType["User"] = 1] = "User";
	    PrincipalType[PrincipalType["DistributionList"] = 2] = "DistributionList";
	    PrincipalType[PrincipalType["SecurityGroup"] = 4] = "SecurityGroup";
	    PrincipalType[PrincipalType["SharePointGroup"] = 8] = "SharePointGroup";
	    PrincipalType[PrincipalType["All"] = 15] = "All";
	})(PrincipalType = exports.PrincipalType || (exports.PrincipalType = {}));
	/**
	 * Describes a collection of site users
	 *
	 */
	var SiteGroups = (function (_super) {
	    __extends(SiteGroups, _super);
	    /**
	     * Creates a new instance of the SiteUsers class
	     *
	     * @param baseUrl The url or Queryable which forms the parent of this user collection
	     */
	    function SiteGroups(baseUrl, path) {
	        if (path === void 0) { path = "sitegroups"; }
	        return _super.call(this, baseUrl, path) || this;
	    }
	    /**
	     * Adds a new group to the site collection
	     *
	     * @param props The properties to be updated
	     */
	    SiteGroups.prototype.add = function (properties) {
	        var _this = this;
	        var postBody = JSON.stringify(util_1.Util.extend({ "__metadata": { "type": "SP.Group" } }, properties));
	        return this.post({ body: postBody }).then(function (data) {
	            return {
	                data: data,
	                group: _this.getById(data.Id),
	            };
	        });
	    };
	    /**
	     * Gets a group from the collection by name
	     *
	     * @param email The name of the group
	     */
	    SiteGroups.prototype.getByName = function (groupName) {
	        return new SiteGroup(this, "getByName('" + groupName + "')");
	    };
	    /**
	     * Gets a group from the collection by id
	     *
	     * @param id The id of the group
	     */
	    SiteGroups.prototype.getById = function (id) {
	        var sg = new SiteGroup(this);
	        sg.concat("(" + id + ")");
	        return sg;
	    };
	    /**
	     * Removes the group with the specified member ID from the collection.
	     *
	     * @param id The id of the group to remove
	     */
	    SiteGroups.prototype.removeById = function (id) {
	        return this.clone(SiteGroups, "removeById('" + id + "')", true).post();
	    };
	    /**
	     * Removes a user from the collection by login name
	     *
	     * @param loginName The login name of the user
	     */
	    SiteGroups.prototype.removeByLoginName = function (loginName) {
	        return this.clone(SiteGroups, "removeByLoginName('" + loginName + "')", true).post();
	    };
	    return SiteGroups;
	}(queryable_1.QueryableCollection));
	exports.SiteGroups = SiteGroups;
	/**
	 * Describes a single group
	 *
	 */
	var SiteGroup = (function (_super) {
	    __extends(SiteGroup, _super);
	    function SiteGroup() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    Object.defineProperty(SiteGroup.prototype, "users", {
	        /**
	         * Get's the users for this group
	         *
	         */
	        get: function () {
	            return new siteusers_1.SiteUsers(this, "users");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    /**
	    * Updates this group instance with the supplied properties
	    *
	    * @param properties A GroupWriteableProperties object of property names and values to update for the user
	    */
	    /* tslint:disable no-string-literal */
	    SiteGroup.prototype.update = function (properties) {
	        var _this = this;
	        var postBody = util_1.Util.extend({ "__metadata": { "type": "SP.Group" } }, properties);
	        return this.post({
	            body: JSON.stringify(postBody),
	            headers: {
	                "X-HTTP-Method": "MERGE",
	            },
	        }).then(function (data) {
	            var retGroup = _this;
	            if (properties.hasOwnProperty("Title")) {
	                retGroup = _this.getParent(SiteGroup, _this.parentUrl, "getByName('" + properties["Title"] + "')");
	            }
	            return {
	                data: data,
	                group: retGroup,
	            };
	        });
	    };
	    return SiteGroup;
	}(queryable_1.QueryableInstance));
	exports.SiteGroup = SiteGroup;


/***/ },
/* 38 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	var __extends = (this && this.__extends) || (function () {
	    var extendStatics = Object.setPrototypeOf ||
	        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
	        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
	    return function (d, b) {
	        extendStatics(d, b);
	        function __() { this.constructor = d; }
	        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
	    };
	})();
	Object.defineProperty(exports, "__esModule", { value: true });
	var queryable_1 = __webpack_require__(22);
	var sitegroups_1 = __webpack_require__(37);
	var util_1 = __webpack_require__(12);
	/**
	 * Describes a collection of all site collection users
	 *
	 */
	var SiteUsers = (function (_super) {
	    __extends(SiteUsers, _super);
	    /**
	     * Creates a new instance of the Users class
	     *
	     * @param baseUrl The url or Queryable which forms the parent of this user collection
	     */
	    function SiteUsers(baseUrl, path) {
	        if (path === void 0) { path = "siteusers"; }
	        return _super.call(this, baseUrl, path) || this;
	    }
	    /**
	     * Gets a user from the collection by email
	     *
	     * @param email The email of the user
	     */
	    SiteUsers.prototype.getByEmail = function (email) {
	        return new SiteUser(this, "getByEmail('" + email + "')");
	    };
	    /**
	     * Gets a user from the collection by id
	     *
	     * @param id The id of the user
	     */
	    SiteUsers.prototype.getById = function (id) {
	        return new SiteUser(this, "getById(" + id + ")");
	    };
	    /**
	     * Gets a user from the collection by login name
	     *
	     * @param loginName The email address of the user
	     */
	    SiteUsers.prototype.getByLoginName = function (loginName) {
	        var su = new SiteUser(this);
	        su.concat("(@v)");
	        su.query.add("@v", "'" + encodeURIComponent(loginName) + "'");
	        return su;
	    };
	    /**
	     * Removes a user from the collection by id
	     *
	     * @param id The id of the user
	     */
	    SiteUsers.prototype.removeById = function (id) {
	        return this.clone(SiteUsers, "removeById(" + id + ")", true).post();
	    };
	    /**
	     * Removes a user from the collection by login name
	     *
	     * @param loginName The login name of the user
	     */
	    SiteUsers.prototype.removeByLoginName = function (loginName) {
	        var o = this.clone(SiteUsers, "removeByLoginName(@v)", true);
	        o.query.add("@v", "'" + encodeURIComponent(loginName) + "'");
	        return o.post();
	    };
	    /**
	     * Add a user to a group
	     *
	     * @param loginName The login name of the user to add to the group
	     *
	     */
	    SiteUsers.prototype.add = function (loginName) {
	        var _this = this;
	        return this.clone(SiteUsers, null, true).post({
	            body: JSON.stringify({ "__metadata": { "type": "SP.User" }, LoginName: loginName }),
	        }).then(function () { return _this.getByLoginName(loginName); });
	    };
	    return SiteUsers;
	}(queryable_1.QueryableCollection));
	exports.SiteUsers = SiteUsers;
	/**
	 * Describes a single user
	 *
	 */
	var SiteUser = (function (_super) {
	    __extends(SiteUser, _super);
	    function SiteUser() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    Object.defineProperty(SiteUser.prototype, "groups", {
	        /**
	         * Get's the groups for this user.
	         *
	         */
	        get: function () {
	            return new sitegroups_1.SiteGroups(this, "groups");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    /**
	    * Updates this user instance with the supplied properties
	    *
	    * @param properties A plain object of property names and values to update for the user
	    */
	    SiteUser.prototype.update = function (properties) {
	        var _this = this;
	        var postBody = util_1.Util.extend({ "__metadata": { "type": "SP.User" } }, properties);
	        return this.post({
	            body: JSON.stringify(postBody),
	            headers: {
	                "X-HTTP-Method": "MERGE",
	            },
	        }).then(function (data) {
	            return {
	                data: data,
	                user: _this,
	            };
	        });
	    };
	    /**
	     * Delete this user
	     *
	     */
	    SiteUser.prototype.delete = function () {
	        return this.post({
	            headers: {
	                "X-HTTP-Method": "DELETE",
	            },
	        });
	    };
	    return SiteUser;
	}(queryable_1.QueryableInstance));
	exports.SiteUser = SiteUser;
	/**
	 * Represents the current user
	 */
	var CurrentUser = (function (_super) {
	    __extends(CurrentUser, _super);
	    function CurrentUser(baseUrl, path) {
	        if (path === void 0) { path = "currentuser"; }
	        return _super.call(this, baseUrl, path) || this;
	    }
	    return CurrentUser;
	}(queryable_1.QueryableInstance));
	exports.CurrentUser = CurrentUser;


/***/ },
/* 39 */
/***/ function(module, exports) {

	// reference: https://msdn.microsoft.com/en-us/library/office/dn600183.aspx
	"use strict";
	Object.defineProperty(exports, "__esModule", { value: true });
	/**
	 * Determines the display mode of the given control or view
	 */
	var ControlMode;
	(function (ControlMode) {
	    ControlMode[ControlMode["Display"] = 1] = "Display";
	    ControlMode[ControlMode["Edit"] = 2] = "Edit";
	    ControlMode[ControlMode["New"] = 3] = "New";
	})(ControlMode = exports.ControlMode || (exports.ControlMode = {}));
	/**
	 * Specifies the type of the field.
	 */
	var FieldTypes;
	(function (FieldTypes) {
	    FieldTypes[FieldTypes["Invalid"] = 0] = "Invalid";
	    FieldTypes[FieldTypes["Integer"] = 1] = "Integer";
	    FieldTypes[FieldTypes["Text"] = 2] = "Text";
	    FieldTypes[FieldTypes["Note"] = 3] = "Note";
	    FieldTypes[FieldTypes["DateTime"] = 4] = "DateTime";
	    FieldTypes[FieldTypes["Counter"] = 5] = "Counter";
	    FieldTypes[FieldTypes["Choice"] = 6] = "Choice";
	    FieldTypes[FieldTypes["Lookup"] = 7] = "Lookup";
	    FieldTypes[FieldTypes["Boolean"] = 8] = "Boolean";
	    FieldTypes[FieldTypes["Number"] = 9] = "Number";
	    FieldTypes[FieldTypes["Currency"] = 10] = "Currency";
	    FieldTypes[FieldTypes["URL"] = 11] = "URL";
	    FieldTypes[FieldTypes["Computed"] = 12] = "Computed";
	    FieldTypes[FieldTypes["Threading"] = 13] = "Threading";
	    FieldTypes[FieldTypes["Guid"] = 14] = "Guid";
	    FieldTypes[FieldTypes["MultiChoice"] = 15] = "MultiChoice";
	    FieldTypes[FieldTypes["GridChoice"] = 16] = "GridChoice";
	    FieldTypes[FieldTypes["Calculated"] = 17] = "Calculated";
	    FieldTypes[FieldTypes["File"] = 18] = "File";
	    FieldTypes[FieldTypes["Attachments"] = 19] = "Attachments";
	    FieldTypes[FieldTypes["User"] = 20] = "User";
	    FieldTypes[FieldTypes["Recurrence"] = 21] = "Recurrence";
	    FieldTypes[FieldTypes["CrossProjectLink"] = 22] = "CrossProjectLink";
	    FieldTypes[FieldTypes["ModStat"] = 23] = "ModStat";
	    FieldTypes[FieldTypes["Error"] = 24] = "Error";
	    FieldTypes[FieldTypes["ContentTypeId"] = 25] = "ContentTypeId";
	    FieldTypes[FieldTypes["PageSeparator"] = 26] = "PageSeparator";
	    FieldTypes[FieldTypes["ThreadIndex"] = 27] = "ThreadIndex";
	    FieldTypes[FieldTypes["WorkflowStatus"] = 28] = "WorkflowStatus";
	    FieldTypes[FieldTypes["AllDayEvent"] = 29] = "AllDayEvent";
	    FieldTypes[FieldTypes["WorkflowEventType"] = 30] = "WorkflowEventType";
	})(FieldTypes = exports.FieldTypes || (exports.FieldTypes = {}));
	var DateTimeFieldFormatType;
	(function (DateTimeFieldFormatType) {
	    DateTimeFieldFormatType[DateTimeFieldFormatType["DateOnly"] = 0] = "DateOnly";
	    DateTimeFieldFormatType[DateTimeFieldFormatType["DateTime"] = 1] = "DateTime";
	})(DateTimeFieldFormatType = exports.DateTimeFieldFormatType || (exports.DateTimeFieldFormatType = {}));
	/**
	 * Specifies the control settings while adding a field.
	 */
	var AddFieldOptions;
	(function (AddFieldOptions) {
	    /**
	     *  Specify that a new field added to the list must also be added to the default content type in the site collection
	     */
	    AddFieldOptions[AddFieldOptions["DefaultValue"] = 0] = "DefaultValue";
	    /**
	     * Specify that a new field added to the list must also be added to the default content type in the site collection.
	     */
	    AddFieldOptions[AddFieldOptions["AddToDefaultContentType"] = 1] = "AddToDefaultContentType";
	    /**
	     * Specify that a new field must not be added to any other content type
	     */
	    AddFieldOptions[AddFieldOptions["AddToNoContentType"] = 2] = "AddToNoContentType";
	    /**
	     *  Specify that a new field that is added to the specified list must also be added to all content types in the site collection
	     */
	    AddFieldOptions[AddFieldOptions["AddToAllContentTypes"] = 4] = "AddToAllContentTypes";
	    /**
	     * Specify adding an internal field name hint for the purpose of avoiding possible database locking or field renaming operations
	     */
	    AddFieldOptions[AddFieldOptions["AddFieldInternalNameHint"] = 8] = "AddFieldInternalNameHint";
	    /**
	     * Specify that a new field that is added to the specified list must also be added to the default list view
	     */
	    AddFieldOptions[AddFieldOptions["AddFieldToDefaultView"] = 16] = "AddFieldToDefaultView";
	    /**
	     * Specify to confirm that no other field has the same display name
	     */
	    AddFieldOptions[AddFieldOptions["AddFieldCheckDisplayName"] = 32] = "AddFieldCheckDisplayName";
	})(AddFieldOptions = exports.AddFieldOptions || (exports.AddFieldOptions = {}));
	var CalendarType;
	(function (CalendarType) {
	    CalendarType[CalendarType["Gregorian"] = 1] = "Gregorian";
	    CalendarType[CalendarType["Japan"] = 3] = "Japan";
	    CalendarType[CalendarType["Taiwan"] = 4] = "Taiwan";
	    CalendarType[CalendarType["Korea"] = 5] = "Korea";
	    CalendarType[CalendarType["Hijri"] = 6] = "Hijri";
	    CalendarType[CalendarType["Thai"] = 7] = "Thai";
	    CalendarType[CalendarType["Hebrew"] = 8] = "Hebrew";
	    CalendarType[CalendarType["GregorianMEFrench"] = 9] = "GregorianMEFrench";
	    CalendarType[CalendarType["GregorianArabic"] = 10] = "GregorianArabic";
	    CalendarType[CalendarType["GregorianXLITEnglish"] = 11] = "GregorianXLITEnglish";
	    CalendarType[CalendarType["GregorianXLITFrench"] = 12] = "GregorianXLITFrench";
	    CalendarType[CalendarType["KoreaJapanLunar"] = 14] = "KoreaJapanLunar";
	    CalendarType[CalendarType["ChineseLunar"] = 15] = "ChineseLunar";
	    CalendarType[CalendarType["SakaEra"] = 16] = "SakaEra";
	    CalendarType[CalendarType["UmAlQura"] = 23] = "UmAlQura";
	})(CalendarType = exports.CalendarType || (exports.CalendarType = {}));
	var UrlFieldFormatType;
	(function (UrlFieldFormatType) {
	    UrlFieldFormatType[UrlFieldFormatType["Hyperlink"] = 0] = "Hyperlink";
	    UrlFieldFormatType[UrlFieldFormatType["Image"] = 1] = "Image";
	})(UrlFieldFormatType = exports.UrlFieldFormatType || (exports.UrlFieldFormatType = {}));
	var PermissionKind;
	(function (PermissionKind) {
	    /**
	     * Has no permissions on the Site. Not available through the user interface.
	     */
	    PermissionKind[PermissionKind["EmptyMask"] = 0] = "EmptyMask";
	    /**
	     * View items in lists, documents in document libraries, and Web discussion comments.
	     */
	    PermissionKind[PermissionKind["ViewListItems"] = 1] = "ViewListItems";
	    /**
	     * Add items to lists, documents to document libraries, and Web discussion comments.
	     */
	    PermissionKind[PermissionKind["AddListItems"] = 2] = "AddListItems";
	    /**
	     * Edit items in lists, edit documents in document libraries, edit Web discussion comments
	     * in documents, and customize Web Part Pages in document libraries.
	     */
	    PermissionKind[PermissionKind["EditListItems"] = 3] = "EditListItems";
	    /**
	     * Delete items from a list, documents from a document library, and Web discussion
	     * comments in documents.
	     */
	    PermissionKind[PermissionKind["DeleteListItems"] = 4] = "DeleteListItems";
	    /**
	     * Approve a minor version of a list item or document.
	     */
	    PermissionKind[PermissionKind["ApproveItems"] = 5] = "ApproveItems";
	    /**
	     * View the source of documents with server-side file handlers.
	     */
	    PermissionKind[PermissionKind["OpenItems"] = 6] = "OpenItems";
	    /**
	     * View past versions of a list item or document.
	     */
	    PermissionKind[PermissionKind["ViewVersions"] = 7] = "ViewVersions";
	    /**
	     * Delete past versions of a list item or document.
	     */
	    PermissionKind[PermissionKind["DeleteVersions"] = 8] = "DeleteVersions";
	    /**
	     * Discard or check in a document which is checked out to another user.
	     */
	    PermissionKind[PermissionKind["CancelCheckout"] = 9] = "CancelCheckout";
	    /**
	     * Create, change, and delete personal views of lists.
	     */
	    PermissionKind[PermissionKind["ManagePersonalViews"] = 10] = "ManagePersonalViews";
	    /**
	     * Create and delete lists, add or remove columns in a list, and add or remove public views of a list.
	     */
	    PermissionKind[PermissionKind["ManageLists"] = 12] = "ManageLists";
	    /**
	     * View forms, views, and application pages, and enumerate lists.
	     */
	    PermissionKind[PermissionKind["ViewFormPages"] = 13] = "ViewFormPages";
	    /**
	     * Make content of a list or document library retrieveable for anonymous users through SharePoint search.
	     * The list permissions in the site do not change.
	     */
	    PermissionKind[PermissionKind["AnonymousSearchAccessList"] = 14] = "AnonymousSearchAccessList";
	    /**
	     * Allow users to open a Site, list, or folder to access items inside that container.
	     */
	    PermissionKind[PermissionKind["Open"] = 17] = "Open";
	    /**
	     * View pages in a Site.
	     */
	    PermissionKind[PermissionKind["ViewPages"] = 18] = "ViewPages";
	    /**
	     * Add, change, or delete HTML pages or Web Part Pages, and edit the Site using
	     * a Windows SharePoint Services compatible editor.
	     */
	    PermissionKind[PermissionKind["AddAndCustomizePages"] = 19] = "AddAndCustomizePages";
	    /**
	     * Apply a theme or borders to the entire Site.
	     */
	    PermissionKind[PermissionKind["ApplyThemeAndBorder"] = 20] = "ApplyThemeAndBorder";
	    /**
	     * Apply a style sheet (.css file) to the Site.
	     */
	    PermissionKind[PermissionKind["ApplyStyleSheets"] = 21] = "ApplyStyleSheets";
	    /**
	     * View reports on Site usage.
	     */
	    PermissionKind[PermissionKind["ViewUsageData"] = 22] = "ViewUsageData";
	    /**
	     * Create a Site using Self-Service Site Creation.
	     */
	    PermissionKind[PermissionKind["CreateSSCSite"] = 23] = "CreateSSCSite";
	    /**
	     * Create subsites such as team sites, Meeting Workspace sites, and Document Workspace sites.
	     */
	    PermissionKind[PermissionKind["ManageSubwebs"] = 24] = "ManageSubwebs";
	    /**
	     * Create a group of users that can be used anywhere within the site collection.
	     */
	    PermissionKind[PermissionKind["CreateGroups"] = 25] = "CreateGroups";
	    /**
	     * Create and change permission levels on the Site and assign permissions to users
	     * and groups.
	     */
	    PermissionKind[PermissionKind["ManagePermissions"] = 26] = "ManagePermissions";
	    /**
	     * Enumerate files and folders in a Site using Microsoft Office SharePoint Designer
	     * and WebDAV interfaces.
	     */
	    PermissionKind[PermissionKind["BrowseDirectories"] = 27] = "BrowseDirectories";
	    /**
	     * View information about users of the Site.
	     */
	    PermissionKind[PermissionKind["BrowseUserInfo"] = 28] = "BrowseUserInfo";
	    /**
	     * Add or remove personal Web Parts on a Web Part Page.
	     */
	    PermissionKind[PermissionKind["AddDelPrivateWebParts"] = 29] = "AddDelPrivateWebParts";
	    /**
	     * Update Web Parts to display personalized information.
	     */
	    PermissionKind[PermissionKind["UpdatePersonalWebParts"] = 30] = "UpdatePersonalWebParts";
	    /**
	     * Grant the ability to perform all administration tasks for the Site as well as
	     * manage content, activate, deactivate, or edit properties of Site scoped Features
	     * through the object model or through the user interface (UI). When granted on the
	     * root Site of a Site Collection, activate, deactivate, or edit properties of
	     * site collection scoped Features through the object model. To browse to the Site
	     * Collection Features page and activate or deactivate Site Collection scoped Features
	     * through the UI, you must be a Site Collection administrator.
	     */
	    PermissionKind[PermissionKind["ManageWeb"] = 31] = "ManageWeb";
	    /**
	     * Content of lists and document libraries in the Web site will be retrieveable for anonymous users through
	     * SharePoint search if the list or document library has AnonymousSearchAccessList set.
	     */
	    PermissionKind[PermissionKind["AnonymousSearchAccessWebLists"] = 32] = "AnonymousSearchAccessWebLists";
	    /**
	     * Use features that launch client applications. Otherwise, users must work on documents
	     * locally and upload changes.
	     */
	    PermissionKind[PermissionKind["UseClientIntegration"] = 37] = "UseClientIntegration";
	    /**
	     * Use SOAP, WebDAV, or Microsoft Office SharePoint Designer interfaces to access the Site.
	     */
	    PermissionKind[PermissionKind["UseRemoteAPIs"] = 38] = "UseRemoteAPIs";
	    /**
	     * Manage alerts for all users of the Site.
	     */
	    PermissionKind[PermissionKind["ManageAlerts"] = 39] = "ManageAlerts";
	    /**
	     * Create e-mail alerts.
	     */
	    PermissionKind[PermissionKind["CreateAlerts"] = 40] = "CreateAlerts";
	    /**
	     * Allows a user to change his or her user information, such as adding a picture.
	     */
	    PermissionKind[PermissionKind["EditMyUserInfo"] = 41] = "EditMyUserInfo";
	    /**
	     * Enumerate permissions on Site, list, folder, document, or list item.
	     */
	    PermissionKind[PermissionKind["EnumeratePermissions"] = 63] = "EnumeratePermissions";
	    /**
	     * Has all permissions on the Site. Not available through the user interface.
	     */
	    PermissionKind[PermissionKind["FullMask"] = 65] = "FullMask";
	})(PermissionKind = exports.PermissionKind || (exports.PermissionKind = {}));
	var PrincipalType;
	(function (PrincipalType) {
	    PrincipalType[PrincipalType["None"] = 0] = "None";
	    PrincipalType[PrincipalType["User"] = 1] = "User";
	    PrincipalType[PrincipalType["DistributionList"] = 2] = "DistributionList";
	    PrincipalType[PrincipalType["SecurityGroup"] = 4] = "SecurityGroup";
	    PrincipalType[PrincipalType["SharePointGroup"] = 8] = "SharePointGroup";
	    PrincipalType[PrincipalType["All"] = 15] = "All";
	})(PrincipalType = exports.PrincipalType || (exports.PrincipalType = {}));
	var RoleType;
	(function (RoleType) {
	    RoleType[RoleType["None"] = 0] = "None";
	    RoleType[RoleType["Guest"] = 1] = "Guest";
	    RoleType[RoleType["Reader"] = 2] = "Reader";
	    RoleType[RoleType["Contributor"] = 3] = "Contributor";
	    RoleType[RoleType["WebDesigner"] = 4] = "WebDesigner";
	    RoleType[RoleType["Administrator"] = 5] = "Administrator";
	})(RoleType = exports.RoleType || (exports.RoleType = {}));
	var PageType;
	(function (PageType) {
	    PageType[PageType["Invalid"] = -1] = "Invalid";
	    PageType[PageType["DefaultView"] = 0] = "DefaultView";
	    PageType[PageType["NormalView"] = 1] = "NormalView";
	    PageType[PageType["DialogView"] = 2] = "DialogView";
	    PageType[PageType["View"] = 3] = "View";
	    PageType[PageType["DisplayForm"] = 4] = "DisplayForm";
	    PageType[PageType["DisplayFormDialog"] = 5] = "DisplayFormDialog";
	    PageType[PageType["EditForm"] = 6] = "EditForm";
	    PageType[PageType["EditFormDialog"] = 7] = "EditFormDialog";
	    PageType[PageType["NewForm"] = 8] = "NewForm";
	    PageType[PageType["NewFormDialog"] = 9] = "NewFormDialog";
	    PageType[PageType["SolutionForm"] = 10] = "SolutionForm";
	    PageType[PageType["PAGE_MAXITEMS"] = 11] = "PAGE_MAXITEMS";
	})(PageType = exports.PageType || (exports.PageType = {}));
	var SharingLinkKind;
	(function (SharingLinkKind) {
	    /**
	     * Uninitialized link
	     */
	    SharingLinkKind[SharingLinkKind["Uninitialized"] = 0] = "Uninitialized";
	    /**
	     * Direct link to the object being shared
	     */
	    SharingLinkKind[SharingLinkKind["Direct"] = 1] = "Direct";
	    /**
	     * Organization-shareable link to the object being shared with view permissions
	     */
	    SharingLinkKind[SharingLinkKind["OrganizationView"] = 2] = "OrganizationView";
	    /**
	     * Organization-shareable link to the object being shared with edit permissions
	     */
	    SharingLinkKind[SharingLinkKind["OrganizationEdit"] = 3] = "OrganizationEdit";
	    /**
	     * View only anonymous link
	     */
	    SharingLinkKind[SharingLinkKind["AnonymousView"] = 4] = "AnonymousView";
	    /**
	     * Read/Write anonymous link
	     */
	    SharingLinkKind[SharingLinkKind["AnonymousEdit"] = 5] = "AnonymousEdit";
	    /**
	     * Flexible sharing Link where properties can change without affecting link URL
	     */
	    SharingLinkKind[SharingLinkKind["Flexible"] = 6] = "Flexible";
	})(SharingLinkKind = exports.SharingLinkKind || (exports.SharingLinkKind = {}));
	;
	/**
	 * Indicates the role of the sharing link
	 */
	var SharingRole;
	(function (SharingRole) {
	    SharingRole[SharingRole["None"] = 0] = "None";
	    SharingRole[SharingRole["View"] = 1] = "View";
	    SharingRole[SharingRole["Edit"] = 2] = "Edit";
	    SharingRole[SharingRole["Owner"] = 3] = "Owner";
	})(SharingRole = exports.SharingRole || (exports.SharingRole = {}));
	var SharingOperationStatusCode;
	(function (SharingOperationStatusCode) {
	    /**
	     * The share operation completed without errors.
	     */
	    SharingOperationStatusCode[SharingOperationStatusCode["CompletedSuccessfully"] = 0] = "CompletedSuccessfully";
	    /**
	     * The share operation completed and generated requests for access.
	     */
	    SharingOperationStatusCode[SharingOperationStatusCode["AccessRequestsQueued"] = 1] = "AccessRequestsQueued";
	    /**
	     * The share operation failed as there were no resolved users.
	     */
	    SharingOperationStatusCode[SharingOperationStatusCode["NoResolvedUsers"] = -1] = "NoResolvedUsers";
	    /**
	     * The share operation failed due to insufficient permissions.
	     */
	    SharingOperationStatusCode[SharingOperationStatusCode["AccessDenied"] = -2] = "AccessDenied";
	    /**
	     * The share operation failed when attempting a cross site share, which is not supported.
	     */
	    SharingOperationStatusCode[SharingOperationStatusCode["CrossSiteRequestNotSupported"] = -3] = "CrossSiteRequestNotSupported";
	    /**
	     * The sharing operation failed due to an unknown error.
	     */
	    SharingOperationStatusCode[SharingOperationStatusCode["UnknowError"] = -4] = "UnknowError";
	    /**
	     * The text you typed is too long. Please shorten it.
	     */
	    SharingOperationStatusCode[SharingOperationStatusCode["EmailBodyTooLong"] = -5] = "EmailBodyTooLong";
	    /**
	     * The maximum number of unique scopes in the list has been exceeded.
	     */
	    SharingOperationStatusCode[SharingOperationStatusCode["ListUniqueScopesExceeded"] = -6] = "ListUniqueScopesExceeded";
	    /**
	     * The share operation failed because a sharing capability is disabled in the site.
	     */
	    SharingOperationStatusCode[SharingOperationStatusCode["CapabilityDisabled"] = -7] = "CapabilityDisabled";
	    /**
	     * The specified object for the share operation is not supported.
	     */
	    SharingOperationStatusCode[SharingOperationStatusCode["ObjectNotSupported"] = -8] = "ObjectNotSupported";
	    /**
	     * A SharePoint group cannot contain another SharePoint group.
	     */
	    SharingOperationStatusCode[SharingOperationStatusCode["NestedGroupsNotSupported"] = -9] = "NestedGroupsNotSupported";
	})(SharingOperationStatusCode = exports.SharingOperationStatusCode || (exports.SharingOperationStatusCode = {}));
	var SPSharedObjectType;
	(function (SPSharedObjectType) {
	    SPSharedObjectType[SPSharedObjectType["Unknown"] = 0] = "Unknown";
	    SPSharedObjectType[SPSharedObjectType["File"] = 1] = "File";
	    SPSharedObjectType[SPSharedObjectType["Folder"] = 2] = "Folder";
	    SPSharedObjectType[SPSharedObjectType["Item"] = 3] = "Item";
	    SPSharedObjectType[SPSharedObjectType["List"] = 4] = "List";
	    SPSharedObjectType[SPSharedObjectType["Web"] = 5] = "Web";
	    SPSharedObjectType[SPSharedObjectType["Max"] = 6] = "Max";
	})(SPSharedObjectType = exports.SPSharedObjectType || (exports.SPSharedObjectType = {}));
	var SharingDomainRestrictionMode;
	(function (SharingDomainRestrictionMode) {
	    SharingDomainRestrictionMode[SharingDomainRestrictionMode["None"] = 0] = "None";
	    SharingDomainRestrictionMode[SharingDomainRestrictionMode["AllowList"] = 1] = "AllowList";
	    SharingDomainRestrictionMode[SharingDomainRestrictionMode["BlockList"] = 2] = "BlockList";
	})(SharingDomainRestrictionMode = exports.SharingDomainRestrictionMode || (exports.SharingDomainRestrictionMode = {}));
	;


/***/ },
/* 40 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	var __extends = (this && this.__extends) || (function () {
	    var extendStatics = Object.setPrototypeOf ||
	        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
	        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
	    return function (d, b) {
	        extendStatics(d, b);
	        function __() { this.constructor = d; }
	        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
	    };
	})();
	Object.defineProperty(exports, "__esModule", { value: true });
	var queryable_1 = __webpack_require__(22);
	var queryableshareable_1 = __webpack_require__(34);
	var files_1 = __webpack_require__(41);
	var util_1 = __webpack_require__(12);
	var odata_1 = __webpack_require__(23);
	var items_1 = __webpack_require__(33);
	/**
	 * Describes a collection of Folder objects
	 *
	 */
	var Folders = (function (_super) {
	    __extends(Folders, _super);
	    /**
	     * Creates a new instance of the Folders class
	     *
	     * @param baseUrl The url or Queryable which forms the parent of this fields collection
	     */
	    function Folders(baseUrl, path) {
	        if (path === void 0) { path = "folders"; }
	        return _super.call(this, baseUrl, path) || this;
	    }
	    /**
	     * Gets a folder by folder name
	     *
	     */
	    Folders.prototype.getByName = function (name) {
	        var f = new Folder(this);
	        f.concat("('" + name + "')");
	        return f;
	    };
	    /**
	     * Adds a new folder to the current folder (relative) or any folder (absolute)
	     *
	     * @param url The relative or absolute url where the new folder will be created. Urls starting with a forward slash are absolute.
	     * @returns The new Folder and the raw response.
	     */
	    Folders.prototype.add = function (url) {
	        var _this = this;
	        return this.clone(Folders, "add('" + url + "')", true).post().then(function (response) {
	            return {
	                data: response,
	                folder: _this.getByName(url),
	            };
	        });
	    };
	    return Folders;
	}(queryable_1.QueryableCollection));
	exports.Folders = Folders;
	/**
	 * Describes a single Folder instance
	 *
	 */
	var Folder = (function (_super) {
	    __extends(Folder, _super);
	    function Folder() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    Object.defineProperty(Folder.prototype, "contentTypeOrder", {
	        /**
	         * Specifies the sequence in which content types are displayed.
	         *
	         */
	        get: function () {
	            return new queryable_1.QueryableCollection(this, "contentTypeOrder");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Folder.prototype, "files", {
	        /**
	         * Gets this folder's files
	         *
	         */
	        get: function () {
	            return new files_1.Files(this);
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Folder.prototype, "folders", {
	        /**
	         * Gets this folder's sub folders
	         *
	         */
	        get: function () {
	            return new Folders(this);
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Folder.prototype, "listItemAllFields", {
	        /**
	         * Gets this folder's list item field values
	         *
	         */
	        get: function () {
	            return new queryable_1.QueryableCollection(this, "listItemAllFields");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Folder.prototype, "parentFolder", {
	        /**
	         * Gets the parent folder, if available
	         *
	         */
	        get: function () {
	            return new Folder(this, "parentFolder");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Folder.prototype, "properties", {
	        /**
	         * Gets this folder's properties
	         *
	         */
	        get: function () {
	            return new queryable_1.QueryableInstance(this, "properties");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Folder.prototype, "serverRelativeUrl", {
	        /**
	         * Gets this folder's server relative url
	         *
	         */
	        get: function () {
	            return new queryable_1.Queryable(this, "serverRelativeUrl");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Folder.prototype, "uniqueContentTypeOrder", {
	        /**
	         * Gets a value that specifies the content type order.
	         *
	         */
	        get: function () {
	            return new queryable_1.QueryableCollection(this, "uniqueContentTypeOrder");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Folder.prototype.update = function (properties) {
	        var _this = this;
	        var postBody = JSON.stringify(util_1.Util.extend({
	            "__metadata": { "type": "SP.Folder" },
	        }, properties));
	        return this.post({
	            body: postBody,
	            headers: {
	                "X-HTTP-Method": "MERGE",
	            },
	        }).then(function (data) {
	            return {
	                data: data,
	                folder: _this,
	            };
	        });
	    };
	    /**
	    * Delete this folder
	    *
	    * @param eTag Value used in the IF-Match header, by default "*"
	    */
	    Folder.prototype.delete = function (eTag) {
	        if (eTag === void 0) { eTag = "*"; }
	        return this.clone(Folder, null, true).post({
	            headers: {
	                "IF-Match": eTag,
	                "X-HTTP-Method": "DELETE",
	            },
	        });
	    };
	    /**
	     * Moves the folder to the Recycle Bin and returns the identifier of the new Recycle Bin item.
	     */
	    Folder.prototype.recycle = function () {
	        return this.clone(Folder, "recycle", true).post();
	    };
	    /**
	     * Gets the associated list item for this folder, loading the default properties
	     */
	    Folder.prototype.getItem = function () {
	        var selects = [];
	        for (var _i = 0; _i < arguments.length; _i++) {
	            selects[_i] = arguments[_i];
	        }
	        var q = this.listItemAllFields;
	        return q.select.apply(q, selects).get().then(function (d) {
	            return util_1.Util.extend(new items_1.Item(odata_1.getEntityUrl(d)), d);
	        });
	    };
	    return Folder;
	}(queryableshareable_1.QueryableShareableFolder));
	exports.Folder = Folder;


/***/ },
/* 41 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	var __extends = (this && this.__extends) || (function () {
	    var extendStatics = Object.setPrototypeOf ||
	        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
	        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
	    return function (d, b) {
	        extendStatics(d, b);
	        function __() { this.constructor = d; }
	        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
	    };
	})();
	Object.defineProperty(exports, "__esModule", { value: true });
	var queryable_1 = __webpack_require__(22);
	var odata_1 = __webpack_require__(23);
	var util_1 = __webpack_require__(12);
	var exceptions_1 = __webpack_require__(26);
	var webparts_1 = __webpack_require__(42);
	var items_1 = __webpack_require__(33);
	var queryableshareable_1 = __webpack_require__(34);
	var odata_2 = __webpack_require__(23);
	/**
	 * Describes a collection of File objects
	 *
	 */
	var Files = (function (_super) {
	    __extends(Files, _super);
	    /**
	     * Creates a new instance of the Files class
	     *
	     * @param baseUrl The url or Queryable which forms the parent of this fields collection
	     */
	    function Files(baseUrl, path) {
	        if (path === void 0) { path = "files"; }
	        return _super.call(this, baseUrl, path) || this;
	    }
	    /**
	     * Gets a File by filename
	     *
	     * @param name The name of the file, including extension.
	     */
	    Files.prototype.getByName = function (name) {
	        var f = new File(this);
	        f.concat("('" + name + "')");
	        return f;
	    };
	    /**
	     * Uploads a file. Not supported for batching
	     *
	     * @param url The folder-relative url of the file.
	     * @param content The file contents blob.
	     * @param shouldOverWrite Should a file with the same name in the same location be overwritten? (default: true)
	     * @returns The new File and the raw response.
	     */
	    Files.prototype.add = function (url, content, shouldOverWrite) {
	        var _this = this;
	        if (shouldOverWrite === void 0) { shouldOverWrite = true; }
	        return new Files(this, "add(overwrite=" + shouldOverWrite + ",url='" + url + "')")
	            .post({
	            body: content,
	        }).then(function (response) {
	            return {
	                data: response,
	                file: _this.getByName(url),
	            };
	        });
	    };
	    /**
	     * Uploads a file. Not supported for batching
	     *
	     * @param url The folder-relative url of the file.
	     * @param content The Blob file content to add
	     * @param progress A callback function which can be used to track the progress of the upload
	     * @param shouldOverWrite Should a file with the same name in the same location be overwritten? (default: true)
	     * @param chunkSize The size of each file slice, in bytes (default: 10485760)
	     * @returns The new File and the raw response.
	     */
	    Files.prototype.addChunked = function (url, content, progress, shouldOverWrite, chunkSize) {
	        var _this = this;
	        if (shouldOverWrite === void 0) { shouldOverWrite = true; }
	        if (chunkSize === void 0) { chunkSize = 10485760; }
	        var adder = this.clone(Files, "add(overwrite=" + shouldOverWrite + ",url='" + url + "')");
	        return adder.post().then(function () { return _this.getByName(url); }).then(function (file) { return file.setContentChunked(content, progress, chunkSize); }).then(function (response) {
	            return {
	                data: response,
	                file: _this.getByName(url),
	            };
	        });
	    };
	    /**
	     * Adds a ghosted file to an existing list or document library. Not supported for batching.
	     *
	     * @param fileUrl The server-relative url where you want to save the file.
	     * @param templateFileType The type of use to create the file.
	     * @returns The template file that was added and the raw response.
	     */
	    Files.prototype.addTemplateFile = function (fileUrl, templateFileType) {
	        var _this = this;
	        return this.clone(Files, "addTemplateFile(urloffile='" + fileUrl + "',templatefiletype=" + templateFileType + ")")
	            .post().then(function (response) {
	            return {
	                data: response,
	                file: _this.getByName(fileUrl),
	            };
	        });
	    };
	    return Files;
	}(queryable_1.QueryableCollection));
	exports.Files = Files;
	/**
	 * Describes a single File instance
	 *
	 */
	var File = (function (_super) {
	    __extends(File, _super);
	    function File() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    Object.defineProperty(File.prototype, "listItemAllFields", {
	        /**
	         * Gets a value that specifies the list item field values for the list item corresponding to the file.
	         *
	         */
	        get: function () {
	            return new queryable_1.QueryableCollection(this, "listItemAllFields");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(File.prototype, "versions", {
	        /**
	         * Gets a collection of versions
	         *
	         */
	        get: function () {
	            return new Versions(this);
	        },
	        enumerable: true,
	        configurable: true
	    });
	    /**
	     * Approves the file submitted for content approval with the specified comment.
	     * Only documents in lists that are enabled for content approval can be approved.
	     *
	     * @param comment The comment for the approval.
	     */
	    File.prototype.approve = function (comment) {
	        return this.clone(File, "approve(comment='" + comment + "')", true).post();
	    };
	    /**
	     * Stops the chunk upload session without saving the uploaded data. Does not support batching.
	     * If the file doesn’t already exist in the library, the partially uploaded file will be deleted.
	     * Use this in response to user action (as in a request to cancel an upload) or an error or exception.
	     * Use the uploadId value that was passed to the StartUpload method that started the upload session.
	     * This method is currently available only on Office 365.
	     *
	     * @param uploadId The unique identifier of the upload session.
	     */
	    File.prototype.cancelUpload = function (uploadId) {
	        return this.clone(File, "cancelUpload(uploadId=guid'" + uploadId + "')", false).post();
	    };
	    /**
	     * Checks the file in to a document library based on the check-in type.
	     *
	     * @param comment A comment for the check-in. Its length must be <= 1023.
	     * @param checkinType The check-in type for the file.
	     */
	    File.prototype.checkin = function (comment, checkinType) {
	        if (comment === void 0) { comment = ""; }
	        if (checkinType === void 0) { checkinType = CheckinType.Major; }
	        if (comment.length > 1023) {
	            throw new exceptions_1.MaxCommentLengthException();
	        }
	        return this.clone(File, "checkin(comment='" + comment + "',checkintype=" + checkinType + ")", true).post();
	    };
	    /**
	     * Checks out the file from a document library.
	     */
	    File.prototype.checkout = function () {
	        return this.clone(File, "checkout", true).post();
	    };
	    /**
	     * Copies the file to the destination url.
	     *
	     * @param url The absolute url or server relative url of the destination file path to copy to.
	     * @param shouldOverWrite Should a file with the same name in the same location be overwritten?
	     */
	    File.prototype.copyTo = function (url, shouldOverWrite) {
	        if (shouldOverWrite === void 0) { shouldOverWrite = true; }
	        return this.clone(File, "copyTo(strnewurl='" + url + "',boverwrite=" + shouldOverWrite + ")", true).post();
	    };
	    /**
	     * Delete this file.
	     *
	     * @param eTag Value used in the IF-Match header, by default "*"
	     */
	    File.prototype.delete = function (eTag) {
	        if (eTag === void 0) { eTag = "*"; }
	        return this.clone(File, null, true).post({
	            headers: {
	                "IF-Match": eTag,
	                "X-HTTP-Method": "DELETE",
	            },
	        });
	    };
	    /**
	     * Denies approval for a file that was submitted for content approval.
	     * Only documents in lists that are enabled for content approval can be denied.
	     *
	     * @param comment The comment for the denial.
	     */
	    File.prototype.deny = function (comment) {
	        if (comment === void 0) { comment = ""; }
	        if (comment.length > 1023) {
	            throw new exceptions_1.MaxCommentLengthException();
	        }
	        return this.clone(File, "deny(comment='" + comment + "')", true).post();
	    };
	    /**
	     * Specifies the control set used to access, modify, or add Web Parts associated with this Web Part Page and view.
	     * An exception is thrown if the file is not an ASPX page.
	     *
	     * @param scope The WebPartsPersonalizationScope view on the Web Parts page.
	     */
	    File.prototype.getLimitedWebPartManager = function (scope) {
	        if (scope === void 0) { scope = WebPartsPersonalizationScope.Shared; }
	        return new webparts_1.LimitedWebPartManager(this, "getLimitedWebPartManager(scope=" + scope + ")");
	    };
	    /**
	     * Moves the file to the specified destination url.
	     *
	     * @param url The absolute url or server relative url of the destination file path to move to.
	     * @param moveOperations The bitwise MoveOperations value for how to move the file.
	     */
	    File.prototype.moveTo = function (url, moveOperations) {
	        if (moveOperations === void 0) { moveOperations = MoveOperations.Overwrite; }
	        return this.clone(File, "moveTo(newurl='" + url + "',flags=" + moveOperations + ")", true).post();
	    };
	    /**
	     * Submits the file for content approval with the specified comment.
	     *
	     * @param comment The comment for the published file. Its length must be <= 1023.
	     */
	    File.prototype.publish = function (comment) {
	        if (comment === void 0) { comment = ""; }
	        if (comment.length > 1023) {
	            throw new exceptions_1.MaxCommentLengthException();
	        }
	        return this.clone(File, "publish(comment='" + comment + "')", true).post();
	    };
	    /**
	     * Moves the file to the Recycle Bin and returns the identifier of the new Recycle Bin item.
	     *
	     * @returns The GUID of the recycled file.
	     */
	    File.prototype.recycle = function () {
	        return this.clone(File, "recycle", true).post();
	    };
	    /**
	     * Reverts an existing checkout for the file.
	     *
	     */
	    File.prototype.undoCheckout = function () {
	        return this.clone(File, "undoCheckout", true).post();
	    };
	    /**
	     * Removes the file from content approval or unpublish a major version.
	     *
	     * @param comment The comment for the unpublish operation. Its length must be <= 1023.
	     */
	    File.prototype.unpublish = function (comment) {
	        if (comment === void 0) { comment = ""; }
	        if (comment.length > 1023) {
	            throw new exceptions_1.MaxCommentLengthException();
	        }
	        return this.clone(File, "unpublish(comment='" + comment + "')", true).post();
	    };
	    /**
	     * Gets the contents of the file as text. Not supported in batching.
	     *
	     */
	    File.prototype.getText = function () {
	        return this.clone(File, "$value").get(new odata_1.TextFileParser(), { headers: { "binaryStringResponseBody": "true" } });
	    };
	    /**
	     * Gets the contents of the file as a blob, does not work in Node.js. Not supported in batching.
	     *
	     */
	    File.prototype.getBlob = function () {
	        return this.clone(File, "$value").get(new odata_1.BlobFileParser(), { headers: { "binaryStringResponseBody": "true" } });
	    };
	    /**
	     * Gets the contents of a file as an ArrayBuffer, works in Node.js. Not supported in batching.
	     */
	    File.prototype.getBuffer = function () {
	        return this.clone(File, "$value").get(new odata_1.BufferFileParser(), { headers: { "binaryStringResponseBody": "true" } });
	    };
	    /**
	     * Gets the contents of a file as an ArrayBuffer, works in Node.js. Not supported in batching.
	     */
	    File.prototype.getJSON = function () {
	        return this.clone(File, "$value").get(new odata_1.JSONFileParser(), { headers: { "binaryStringResponseBody": "true" } });
	    };
	    /**
	     * Sets the content of a file, for large files use setContentChunked. Not supported in batching.
	     *
	     * @param content The file content
	     *
	     */
	    File.prototype.setContent = function (content) {
	        var _this = this;
	        return this.clone(File, "$value").post({
	            body: content,
	            headers: {
	                "X-HTTP-Method": "PUT",
	            },
	        }).then(function (_) { return new File(_this); });
	    };
	    /**
	     * Gets the associated list item for this folder, loading the default properties
	     */
	    File.prototype.getItem = function () {
	        var selects = [];
	        for (var _i = 0; _i < arguments.length; _i++) {
	            selects[_i] = arguments[_i];
	        }
	        var q = this.listItemAllFields;
	        return q.select.apply(q, selects).get().then(function (d) {
	            return util_1.Util.extend(new items_1.Item(odata_2.getEntityUrl(d)), d);
	        });
	    };
	    /**
	     * Sets the contents of a file using a chunked upload approach. Not supported in batching.
	     *
	     * @param file The file to upload
	     * @param progress A callback function which can be used to track the progress of the upload
	     * @param chunkSize The size of each file slice, in bytes (default: 10485760)
	     */
	    File.prototype.setContentChunked = function (file, progress, chunkSize) {
	        if (chunkSize === void 0) { chunkSize = 10485760; }
	        if (typeof progress === "undefined") {
	            progress = function () { return null; };
	        }
	        var self = this;
	        var fileSize = file.size;
	        var blockCount = parseInt((file.size / chunkSize).toString(), 10) + ((file.size % chunkSize === 0) ? 1 : 0);
	        var uploadId = util_1.Util.getGUID();
	        // start the chain with the first fragment
	        progress({ blockNumber: 1, chunkSize: chunkSize, currentPointer: 0, fileSize: fileSize, stage: "starting", totalBlocks: blockCount });
	        var chain = self.startUpload(uploadId, file.slice(0, chunkSize));
	        var _loop_1 = function (i) {
	            chain = chain.then(function (pointer) {
	                progress({ blockNumber: i, chunkSize: chunkSize, currentPointer: pointer, fileSize: fileSize, stage: "continue", totalBlocks: blockCount });
	                return self.continueUpload(uploadId, pointer, file.slice(pointer, pointer + chunkSize));
	            });
	        };
	        // skip the first and last blocks
	        for (var i = 2; i < blockCount; i++) {
	            _loop_1(i);
	        }
	        return chain.then(function (pointer) {
	            progress({ blockNumber: blockCount, chunkSize: chunkSize, currentPointer: pointer, fileSize: fileSize, stage: "finishing", totalBlocks: blockCount });
	            return self.finishUpload(uploadId, pointer, file.slice(pointer));
	        }).then(function (_) {
	            return self;
	        });
	    };
	    /**
	     * Starts a new chunk upload session and uploads the first fragment.
	     * The current file content is not changed when this method completes.
	     * The method is idempotent (and therefore does not change the result) as long as you use the same values for uploadId and stream.
	     * The upload session ends either when you use the CancelUpload method or when you successfully
	     * complete the upload session by passing the rest of the file contents through the ContinueUpload and FinishUpload methods.
	     * The StartUpload and ContinueUpload methods return the size of the running total of uploaded data in bytes,
	     * so you can pass those return values to subsequent uses of ContinueUpload and FinishUpload.
	     * This method is currently available only on Office 365.
	     *
	     * @param uploadId The unique identifier of the upload session.
	     * @param fragment The file contents.
	     * @returns The size of the total uploaded data in bytes.
	     */
	    File.prototype.startUpload = function (uploadId, fragment) {
	        return this.clone(File, "startUpload(uploadId=guid'" + uploadId + "')").postAs({ body: fragment }).then(function (n) { return parseFloat(n); });
	    };
	    /**
	     * Continues the chunk upload session with an additional fragment.
	     * The current file content is not changed.
	     * Use the uploadId value that was passed to the StartUpload method that started the upload session.
	     * This method is currently available only on Office 365.
	     *
	     * @param uploadId The unique identifier of the upload session.
	     * @param fileOffset The size of the offset into the file where the fragment starts.
	     * @param fragment The file contents.
	     * @returns The size of the total uploaded data in bytes.
	     */
	    File.prototype.continueUpload = function (uploadId, fileOffset, fragment) {
	        return this.clone(File, "continueUpload(uploadId=guid'" + uploadId + "',fileOffset=" + fileOffset + ")").postAs({ body: fragment }).then(function (n) { return parseFloat(n); });
	    };
	    /**
	     * Uploads the last file fragment and commits the file. The current file content is changed when this method completes.
	     * Use the uploadId value that was passed to the StartUpload method that started the upload session.
	     * This method is currently available only on Office 365.
	     *
	     * @param uploadId The unique identifier of the upload session.
	     * @param fileOffset The size of the offset into the file where the fragment starts.
	     * @param fragment The file contents.
	     * @returns The newly uploaded file.
	     */
	    File.prototype.finishUpload = function (uploadId, fileOffset, fragment) {
	        return this.clone(File, "finishUpload(uploadId=guid'" + uploadId + "',fileOffset=" + fileOffset + ")")
	            .postAs({ body: fragment }).then(function (response) {
	            return {
	                data: response,
	                file: new File(response.ServerRelativeUrl),
	            };
	        });
	    };
	    return File;
	}(queryableshareable_1.QueryableShareableFile));
	exports.File = File;
	/**
	 * Describes a collection of Version objects
	 *
	 */
	var Versions = (function (_super) {
	    __extends(Versions, _super);
	    /**
	     * Creates a new instance of the File class
	     *
	     * @param baseUrl The url or Queryable which forms the parent of this fields collection
	     */
	    function Versions(baseUrl, path) {
	        if (path === void 0) { path = "versions"; }
	        return _super.call(this, baseUrl, path) || this;
	    }
	    /**
	     * Gets a version by id
	     *
	     * @param versionId The id of the version to retrieve
	     */
	    Versions.prototype.getById = function (versionId) {
	        var v = new Version(this);
	        v.concat("(" + versionId + ")");
	        return v;
	    };
	    /**
	     * Deletes all the file version objects in the collection.
	     *
	     */
	    Versions.prototype.deleteAll = function () {
	        return new Versions(this, "deleteAll").post();
	    };
	    /**
	     * Deletes the specified version of the file.
	     *
	     * @param versionId The ID of the file version to delete.
	     */
	    Versions.prototype.deleteById = function (versionId) {
	        return this.clone(Versions, "deleteById(vid=" + versionId + ")", true).post();
	    };
	    /**
	     * Deletes the file version object with the specified version label.
	     *
	     * @param label The version label of the file version to delete, for example: 1.2
	     */
	    Versions.prototype.deleteByLabel = function (label) {
	        return this.clone(Versions, "deleteByLabel(versionlabel='" + label + "')", true).post();
	    };
	    /**
	     * Creates a new file version from the file specified by the version label.
	     *
	     * @param label The version label of the file version to restore, for example: 1.2
	     */
	    Versions.prototype.restoreByLabel = function (label) {
	        return this.clone(Versions, "restoreByLabel(versionlabel='" + label + "')", true).post();
	    };
	    return Versions;
	}(queryable_1.QueryableCollection));
	exports.Versions = Versions;
	/**
	 * Describes a single Version instance
	 *
	 */
	var Version = (function (_super) {
	    __extends(Version, _super);
	    function Version() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    /**
	    * Delete a specific version of a file.
	    *
	    * @param eTag Value used in the IF-Match header, by default "*"
	    */
	    Version.prototype.delete = function (eTag) {
	        if (eTag === void 0) { eTag = "*"; }
	        return this.post({
	            headers: {
	                "IF-Match": eTag,
	                "X-HTTP-Method": "DELETE",
	            },
	        });
	    };
	    return Version;
	}(queryable_1.QueryableInstance));
	exports.Version = Version;
	var CheckinType;
	(function (CheckinType) {
	    CheckinType[CheckinType["Minor"] = 0] = "Minor";
	    CheckinType[CheckinType["Major"] = 1] = "Major";
	    CheckinType[CheckinType["Overwrite"] = 2] = "Overwrite";
	})(CheckinType = exports.CheckinType || (exports.CheckinType = {}));
	var WebPartsPersonalizationScope;
	(function (WebPartsPersonalizationScope) {
	    WebPartsPersonalizationScope[WebPartsPersonalizationScope["User"] = 0] = "User";
	    WebPartsPersonalizationScope[WebPartsPersonalizationScope["Shared"] = 1] = "Shared";
	})(WebPartsPersonalizationScope = exports.WebPartsPersonalizationScope || (exports.WebPartsPersonalizationScope = {}));
	var MoveOperations;
	(function (MoveOperations) {
	    MoveOperations[MoveOperations["Overwrite"] = 1] = "Overwrite";
	    MoveOperations[MoveOperations["AllowBrokenThickets"] = 8] = "AllowBrokenThickets";
	})(MoveOperations = exports.MoveOperations || (exports.MoveOperations = {}));
	var TemplateFileType;
	(function (TemplateFileType) {
	    TemplateFileType[TemplateFileType["StandardPage"] = 0] = "StandardPage";
	    TemplateFileType[TemplateFileType["WikiPage"] = 1] = "WikiPage";
	    TemplateFileType[TemplateFileType["FormPage"] = 2] = "FormPage";
	})(TemplateFileType = exports.TemplateFileType || (exports.TemplateFileType = {}));


/***/ },
/* 42 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	var __extends = (this && this.__extends) || (function () {
	    var extendStatics = Object.setPrototypeOf ||
	        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
	        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
	    return function (d, b) {
	        extendStatics(d, b);
	        function __() { this.constructor = d; }
	        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
	    };
	})();
	Object.defineProperty(exports, "__esModule", { value: true });
	var queryable_1 = __webpack_require__(22);
	var LimitedWebPartManager = (function (_super) {
	    __extends(LimitedWebPartManager, _super);
	    function LimitedWebPartManager() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    Object.defineProperty(LimitedWebPartManager.prototype, "webparts", {
	        /**
	         * Gets the set of web part definitions contained by this web part manager
	         *
	         */
	        get: function () {
	            return new WebPartDefinitions(this, "webparts");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    /**
	     * Exports a webpart definition
	     *
	     * @param id the GUID id of the definition to export
	     */
	    LimitedWebPartManager.prototype.export = function (id) {
	        return this.clone(LimitedWebPartManager, "ExportWebPart", true).post({
	            body: JSON.stringify({ webPartId: id }),
	        });
	    };
	    /**
	     * Imports a webpart
	     *
	     * @param xml webpart definition which must be valid XML in the .dwp or .webpart format
	     */
	    LimitedWebPartManager.prototype.import = function (xml) {
	        return this.clone(LimitedWebPartManager, "ImportWebPart", true).post({
	            body: JSON.stringify({ webPartXml: xml }),
	        });
	    };
	    return LimitedWebPartManager;
	}(queryable_1.Queryable));
	exports.LimitedWebPartManager = LimitedWebPartManager;
	var WebPartDefinitions = (function (_super) {
	    __extends(WebPartDefinitions, _super);
	    function WebPartDefinitions() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    /**
	     * Gets a web part definition from the collection by id
	     *
	     * @param id GUID id of the web part definition to get
	     */
	    WebPartDefinitions.prototype.getById = function (id) {
	        return new WebPartDefinition(this, "getbyid('" + id + "')");
	    };
	    return WebPartDefinitions;
	}(queryable_1.QueryableCollection));
	exports.WebPartDefinitions = WebPartDefinitions;
	var WebPartDefinition = (function (_super) {
	    __extends(WebPartDefinition, _super);
	    function WebPartDefinition() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    Object.defineProperty(WebPartDefinition.prototype, "webpart", {
	        /**
	         * Gets the webpart information associated with this definition
	         */
	        get: function () {
	            return new WebPart(this);
	        },
	        enumerable: true,
	        configurable: true
	    });
	    /**
	     * Removes a webpart from a page, all settings will be lost
	     */
	    WebPartDefinition.prototype.delete = function () {
	        return this.clone(WebPartDefinition, "DeleteWebPart", true).post();
	    };
	    return WebPartDefinition;
	}(queryable_1.QueryableInstance));
	exports.WebPartDefinition = WebPartDefinition;
	var WebPart = (function (_super) {
	    __extends(WebPart, _super);
	    /**
	     * Creates a new instance of the WebPart class
	     *
	     * @param baseUrl The url or Queryable which forms the parent of this fields collection
	     * @param path Optional, if supplied will be appended to the supplied baseUrl
	     */
	    function WebPart(baseUrl, path) {
	        if (path === void 0) { path = "webpart"; }
	        return _super.call(this, baseUrl, path) || this;
	    }
	    return WebPart;
	}(queryable_1.QueryableInstance));
	exports.WebPart = WebPart;


/***/ },
/* 43 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	var __extends = (this && this.__extends) || (function () {
	    var extendStatics = Object.setPrototypeOf ||
	        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
	        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
	    return function (d, b) {
	        extendStatics(d, b);
	        function __() { this.constructor = d; }
	        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
	    };
	})();
	Object.defineProperty(exports, "__esModule", { value: true });
	var util_1 = __webpack_require__(12);
	var queryable_1 = __webpack_require__(22);
	/**
	 * Describes a collection of content types
	 *
	 */
	var ContentTypes = (function (_super) {
	    __extends(ContentTypes, _super);
	    /**
	     * Creates a new instance of the ContentTypes class
	     *
	     * @param baseUrl The url or Queryable which forms the parent of this content types collection
	     */
	    function ContentTypes(baseUrl, path) {
	        if (path === void 0) { path = "contenttypes"; }
	        return _super.call(this, baseUrl, path) || this;
	    }
	    /**
	     * Gets a ContentType by content type id
	     */
	    ContentTypes.prototype.getById = function (id) {
	        var ct = new ContentType(this);
	        ct.concat("('" + id + "')");
	        return ct;
	    };
	    /**
	     * Adds an existing contenttype to a content type collection
	     *
	     * @param contentTypeId in the following format, for example: 0x010102
	     */
	    ContentTypes.prototype.addAvailableContentType = function (contentTypeId) {
	        var _this = this;
	        var postBody = JSON.stringify({
	            "contentTypeId": contentTypeId,
	        });
	        return this.clone(ContentTypes, "addAvailableContentType", true).postAs({ body: postBody }).then(function (data) {
	            return {
	                contentType: _this.getById(data.id),
	                data: data,
	            };
	        });
	    };
	    /**
	     * Adds a new content type to the collection
	     *
	     * @param id The desired content type id for the new content type (also determines the parent content type)
	     * @param name The name of the content type
	     * @param description The description of the content type
	     * @param group The group in which to add the content type
	     * @param additionalSettings Any additional settings to provide when creating the content type
	     *
	     */
	    ContentTypes.prototype.add = function (id, name, description, group, additionalSettings) {
	        var _this = this;
	        if (description === void 0) { description = ""; }
	        if (group === void 0) { group = "Custom Content Types"; }
	        if (additionalSettings === void 0) { additionalSettings = {}; }
	        var postBody = JSON.stringify(util_1.Util.extend({
	            "Description": description,
	            "Group": group,
	            "Id": { "StringValue": id },
	            "Name": name,
	            "__metadata": { "type": "SP.ContentType" },
	        }, additionalSettings));
	        return this.post({ body: postBody }).then(function (data) {
	            return { contentType: _this.getById(data.id), data: data };
	        });
	    };
	    return ContentTypes;
	}(queryable_1.QueryableCollection));
	exports.ContentTypes = ContentTypes;
	/**
	 * Describes a single ContentType instance
	 *
	 */
	var ContentType = (function (_super) {
	    __extends(ContentType, _super);
	    function ContentType() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    Object.defineProperty(ContentType.prototype, "fieldLinks", {
	        /**
	         * Gets the column (also known as field) references in the content type.
	        */
	        get: function () {
	            return new FieldLinks(this);
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(ContentType.prototype, "fields", {
	        /**
	         * Gets a value that specifies the collection of fields for the content type.
	         */
	        get: function () {
	            return new queryable_1.QueryableCollection(this, "fields");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(ContentType.prototype, "parent", {
	        /**
	         * Gets the parent content type of the content type.
	         */
	        get: function () {
	            return new ContentType(this, "parent");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(ContentType.prototype, "workflowAssociations", {
	        /**
	         * Gets a value that specifies the collection of workflow associations for the content type.
	         */
	        get: function () {
	            return new queryable_1.QueryableCollection(this, "workflowAssociations");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    /**
	     * Delete this content type
	     */
	    ContentType.prototype.delete = function () {
	        return this.post({
	            headers: {
	                "X-HTTP-Method": "DELETE",
	            },
	        });
	    };
	    return ContentType;
	}(queryable_1.QueryableInstance));
	exports.ContentType = ContentType;
	/**
	 * Represents a collection of field link instances
	 */
	var FieldLinks = (function (_super) {
	    __extends(FieldLinks, _super);
	    /**
	     * Creates a new instance of the ContentType class
	     *
	     * @param baseUrl The url or Queryable which forms the parent of this content type instance
	     */
	    function FieldLinks(baseUrl, path) {
	        if (path === void 0) { path = "fieldlinks"; }
	        return _super.call(this, baseUrl, path) || this;
	    }
	    /**
	     * Gets a FieldLink by GUID id
	     *
	     * @param id The GUID id of the field link
	     */
	    FieldLinks.prototype.getById = function (id) {
	        var fl = new FieldLink(this);
	        fl.concat("(guid'" + id + "')");
	        return fl;
	    };
	    return FieldLinks;
	}(queryable_1.QueryableCollection));
	exports.FieldLinks = FieldLinks;
	/**
	 * Represents a field link instance
	 */
	var FieldLink = (function (_super) {
	    __extends(FieldLink, _super);
	    function FieldLink() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    return FieldLink;
	}(queryable_1.QueryableInstance));
	exports.FieldLink = FieldLink;


/***/ },
/* 44 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	var __extends = (this && this.__extends) || (function () {
	    var extendStatics = Object.setPrototypeOf ||
	        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
	        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
	    return function (d, b) {
	        extendStatics(d, b);
	        function __() { this.constructor = d; }
	        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
	    };
	})();
	Object.defineProperty(exports, "__esModule", { value: true });
	var queryable_1 = __webpack_require__(22);
	var odata_1 = __webpack_require__(23);
	/**
	 * Describes a collection of Item objects
	 *
	 */
	var AttachmentFiles = (function (_super) {
	    __extends(AttachmentFiles, _super);
	    /**
	     * Creates a new instance of the AttachmentFiles class
	     *
	     * @param baseUrl The url or Queryable which forms the parent of this attachments collection
	     */
	    function AttachmentFiles(baseUrl, path) {
	        if (path === void 0) { path = "AttachmentFiles"; }
	        return _super.call(this, baseUrl, path) || this;
	    }
	    /**
	     * Gets a Attachment File by filename
	     *
	     * @param name The name of the file, including extension.
	     */
	    AttachmentFiles.prototype.getByName = function (name) {
	        var f = new AttachmentFile(this);
	        f.concat("('" + name + "')");
	        return f;
	    };
	    /**
	     * Adds a new attachment to the collection. Not supported for batching.
	     *
	     * @param name The name of the file, including extension.
	     * @param content The Base64 file content.
	     */
	    AttachmentFiles.prototype.add = function (name, content) {
	        var _this = this;
	        return this.clone(AttachmentFiles, "add(FileName='" + name + "')").post({
	            body: content,
	        }).then(function (response) {
	            return {
	                data: response,
	                file: _this.getByName(name),
	            };
	        });
	    };
	    /**
	     * Adds mjultiple new attachment to the collection. Not supported for batching.
	     *
	     * @files name The collection of files to add
	     */
	    AttachmentFiles.prototype.addMultiple = function (files) {
	        var _this = this;
	        // add the files in series so we don't get update conflicts
	        return files.reduce(function (chain, file) { return chain.then(function () { return _this.clone(AttachmentFiles, "add(FileName='" + file.name + "')").post({
	            body: file.content,
	        }); }); }, Promise.resolve());
	    };
	    return AttachmentFiles;
	}(queryable_1.QueryableCollection));
	exports.AttachmentFiles = AttachmentFiles;
	/**
	 * Describes a single attachment file instance
	 *
	 */
	var AttachmentFile = (function (_super) {
	    __extends(AttachmentFile, _super);
	    function AttachmentFile() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    /**
	     * Gets the contents of the file as text
	     *
	     */
	    AttachmentFile.prototype.getText = function () {
	        return this.clone(AttachmentFile, "$value").get(new odata_1.TextFileParser());
	    };
	    /**
	     * Gets the contents of the file as a blob, does not work in Node.js
	     *
	     */
	    AttachmentFile.prototype.getBlob = function () {
	        return this.clone(AttachmentFile, "$value").get(new odata_1.BlobFileParser());
	    };
	    /**
	     * Gets the contents of a file as an ArrayBuffer, works in Node.js
	     */
	    AttachmentFile.prototype.getBuffer = function () {
	        return this.clone(AttachmentFile, "$value").get(new odata_1.BufferFileParser());
	    };
	    /**
	     * Gets the contents of a file as an ArrayBuffer, works in Node.js
	     */
	    AttachmentFile.prototype.getJSON = function () {
	        return this.clone(AttachmentFile, "$value").get(new odata_1.JSONFileParser());
	    };
	    /**
	     * Sets the content of a file. Not supported for batching
	     *
	     * @param content The value to set for the file contents
	     */
	    AttachmentFile.prototype.setContent = function (content) {
	        var _this = this;
	        return this.clone(AttachmentFile, "$value").post({
	            body: content,
	            headers: {
	                "X-HTTP-Method": "PUT",
	            },
	        }).then(function (_) { return new AttachmentFile(_this); });
	    };
	    /**
	     * Delete this attachment file
	     *
	     * @param eTag Value used in the IF-Match header, by default "*"
	     */
	    AttachmentFile.prototype.delete = function (eTag) {
	        if (eTag === void 0) { eTag = "*"; }
	        return this.post({
	            headers: {
	                "IF-Match": eTag,
	                "X-HTTP-Method": "DELETE",
	            },
	        });
	    };
	    return AttachmentFile;
	}(queryable_1.QueryableInstance));
	exports.AttachmentFile = AttachmentFile;


/***/ },
/* 45 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	var __extends = (this && this.__extends) || (function () {
	    var extendStatics = Object.setPrototypeOf ||
	        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
	        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
	    return function (d, b) {
	        extendStatics(d, b);
	        function __() { this.constructor = d; }
	        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
	    };
	})();
	Object.defineProperty(exports, "__esModule", { value: true });
	var queryable_1 = __webpack_require__(22);
	var util_1 = __webpack_require__(12);
	/**
	 * Describes the views available in the current context
	 *
	 */
	var Views = (function (_super) {
	    __extends(Views, _super);
	    /**
	     * Creates a new instance of the Views class
	     *
	     * @param baseUrl The url or Queryable which forms the parent of this fields collection
	     */
	    function Views(baseUrl, path) {
	        if (path === void 0) { path = "views"; }
	        return _super.call(this, baseUrl, path) || this;
	    }
	    /**
	     * Gets a view by guid id
	     *
	     * @param id The GUID id of the view
	     */
	    Views.prototype.getById = function (id) {
	        var v = new View(this);
	        v.concat("('" + id + "')");
	        return v;
	    };
	    /**
	     * Gets a view by title (case-sensitive)
	     *
	     * @param title The case-sensitive title of the view
	     */
	    Views.prototype.getByTitle = function (title) {
	        return new View(this, "getByTitle('" + title + "')");
	    };
	    /**
	     * Adds a new view to the collection
	     *
	     * @param title The new views's title
	     * @param personalView True if this is a personal view, otherwise false, default = false
	     * @param additionalSettings Will be passed as part of the view creation body
	     */
	    /*tslint:disable max-line-length */
	    Views.prototype.add = function (title, personalView, additionalSettings) {
	        var _this = this;
	        if (personalView === void 0) { personalView = false; }
	        if (additionalSettings === void 0) { additionalSettings = {}; }
	        var postBody = JSON.stringify(util_1.Util.extend({
	            "PersonalView": personalView,
	            "Title": title,
	            "__metadata": { "type": "SP.View" },
	        }, additionalSettings));
	        return this.clone(Views, null, true).postAs({ body: postBody }).then(function (data) {
	            return {
	                data: data,
	                view: _this.getById(data.Id),
	            };
	        });
	    };
	    return Views;
	}(queryable_1.QueryableCollection));
	exports.Views = Views;
	/**
	 * Describes a single View instance
	 *
	 */
	var View = (function (_super) {
	    __extends(View, _super);
	    function View() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    Object.defineProperty(View.prototype, "fields", {
	        get: function () {
	            return new ViewFields(this);
	        },
	        enumerable: true,
	        configurable: true
	    });
	    /**
	     * Updates this view intance with the supplied properties
	     *
	     * @param properties A plain object hash of values to update for the view
	     */
	    View.prototype.update = function (properties) {
	        var _this = this;
	        var postBody = JSON.stringify(util_1.Util.extend({
	            "__metadata": { "type": "SP.View" },
	        }, properties));
	        return this.post({
	            body: postBody,
	            headers: {
	                "X-HTTP-Method": "MERGE",
	            },
	        }).then(function (data) {
	            return {
	                data: data,
	                view: _this,
	            };
	        });
	    };
	    /**
	     * Delete this view
	     *
	     */
	    View.prototype.delete = function () {
	        return this.post({
	            headers: {
	                "X-HTTP-Method": "DELETE",
	            },
	        });
	    };
	    /**
	     * Returns the list view as HTML.
	     *
	     */
	    View.prototype.renderAsHtml = function () {
	        return this.clone(queryable_1.Queryable, "renderashtml", true).get();
	    };
	    return View;
	}(queryable_1.QueryableInstance));
	exports.View = View;
	var ViewFields = (function (_super) {
	    __extends(ViewFields, _super);
	    function ViewFields(baseUrl, path) {
	        if (path === void 0) { path = "viewfields"; }
	        return _super.call(this, baseUrl, path) || this;
	    }
	    /**
	     * Gets a value that specifies the XML schema that represents the collection.
	     */
	    ViewFields.prototype.getSchemaXml = function () {
	        return this.clone(queryable_1.Queryable, "schemaxml", true).get();
	    };
	    /**
	     * Adds the field with the specified field internal name or display name to the collection.
	     *
	     * @param fieldTitleOrInternalName The case-sensitive internal name or display name of the field to add.
	     */
	    ViewFields.prototype.add = function (fieldTitleOrInternalName) {
	        return this.clone(ViewFields, "addviewfield('" + fieldTitleOrInternalName + "')", true).post();
	    };
	    /**
	     * Moves the field with the specified field internal name to the specified position in the collection.
	     *
	     * @param fieldInternalName The case-sensitive internal name of the field to move.
	     * @param index The zero-based index of the new position for the field.
	     */
	    ViewFields.prototype.move = function (fieldInternalName, index) {
	        return this.clone(ViewFields, "moveviewfieldto", true).post({
	            body: JSON.stringify({ "field": fieldInternalName, "index": index }),
	        });
	    };
	    /**
	     * Removes all the fields from the collection.
	     */
	    ViewFields.prototype.removeAll = function () {
	        return this.clone(ViewFields, "removeallviewfields", true).post();
	    };
	    /**
	     * Removes the field with the specified field internal name from the collection.
	     *
	     * @param fieldInternalName The case-sensitive internal name of the field to remove from the view.
	     */
	    ViewFields.prototype.remove = function (fieldInternalName) {
	        return this.clone(ViewFields, "removeviewfield('" + fieldInternalName + "')", true).post();
	    };
	    return ViewFields;
	}(queryable_1.QueryableCollection));
	exports.ViewFields = ViewFields;


/***/ },
/* 46 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	var __extends = (this && this.__extends) || (function () {
	    var extendStatics = Object.setPrototypeOf ||
	        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
	        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
	    return function (d, b) {
	        extendStatics(d, b);
	        function __() { this.constructor = d; }
	        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
	    };
	})();
	Object.defineProperty(exports, "__esModule", { value: true });
	var queryable_1 = __webpack_require__(22);
	var util_1 = __webpack_require__(12);
	var types_1 = __webpack_require__(39);
	/**
	 * Describes a collection of Field objects
	 *
	 */
	var Fields = (function (_super) {
	    __extends(Fields, _super);
	    /**
	     * Creates a new instance of the Fields class
	     *
	     * @param baseUrl The url or Queryable which forms the parent of this fields collection
	     */
	    function Fields(baseUrl, path) {
	        if (path === void 0) { path = "fields"; }
	        return _super.call(this, baseUrl, path) || this;
	    }
	    /**
	     * Gets a field from the collection by title
	     *
	     * @param title The case-sensitive title of the field
	     */
	    Fields.prototype.getByTitle = function (title) {
	        return new Field(this, "getByTitle('" + title + "')");
	    };
	    /**
	     * Gets a field from the collection by using internal name or title
	     *
	     * @param name The case-sensitive internal name or title of the field
	     */
	    Fields.prototype.getByInternalNameOrTitle = function (name) {
	        return new Field(this, "getByInternalNameOrTitle('" + name + "')");
	    };
	    /**
	     * Gets a list from the collection by guid id
	     *
	     * @param title The Id of the list
	     */
	    Fields.prototype.getById = function (id) {
	        var f = new Field(this);
	        f.concat("('" + id + "')");
	        return f;
	    };
	    /**
	     * Creates a field based on the specified schema
	     */
	    Fields.prototype.createFieldAsXml = function (xml) {
	        var _this = this;
	        var info;
	        if (typeof xml === "string") {
	            info = { SchemaXml: xml };
	        }
	        else {
	            info = xml;
	        }
	        var postBody = JSON.stringify({
	            "parameters": util_1.Util.extend({
	                "__metadata": {
	                    "type": "SP.XmlSchemaFieldCreationInformation",
	                },
	            }, info),
	        });
	        return this.clone(Fields, "createfieldasxml", true).postAs({ body: postBody }).then(function (data) {
	            return {
	                data: data,
	                field: _this.getById(data.Id),
	            };
	        });
	    };
	    /**
	     * Adds a new list to the collection
	     *
	     * @param title The new field's title
	     * @param fieldType The new field's type (ex: SP.FieldText)
	     * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
	     */
	    Fields.prototype.add = function (title, fieldType, properties) {
	        var _this = this;
	        if (properties === void 0) { properties = {}; }
	        var postBody = JSON.stringify(util_1.Util.extend({
	            "Title": title,
	            "__metadata": { "type": fieldType },
	        }, properties));
	        return this.clone(Fields, null, true).postAs({ body: postBody }).then(function (data) {
	            return {
	                data: data,
	                field: _this.getById(data.Id),
	            };
	        });
	    };
	    /**
	     * Adds a new SP.FieldText to the collection
	     *
	     * @param title The field title
	     * @param maxLength The maximum number of characters allowed in the value of the field.
	     * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
	     */
	    Fields.prototype.addText = function (title, maxLength, properties) {
	        if (maxLength === void 0) { maxLength = 255; }
	        var props = {
	            FieldTypeKind: 2,
	            MaxLength: maxLength,
	        };
	        return this.add(title, "SP.FieldText", util_1.Util.extend(props, properties));
	    };
	    /**
	     * Adds a new SP.FieldCalculated to the collection
	     *
	     * @param title The field title.
	     * @param formula The formula for the field.
	     * @param dateFormat The date and time format that is displayed in the field.
	     * @param outputType Specifies the output format for the field. Represents a FieldType value.
	     * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
	     */
	    Fields.prototype.addCalculated = function (title, formula, dateFormat, outputType, properties) {
	        if (outputType === void 0) { outputType = types_1.FieldTypes.Text; }
	        var props = {
	            DateFormat: dateFormat,
	            FieldTypeKind: 17,
	            Formula: formula,
	            OutputType: outputType,
	        };
	        return this.add(title, "SP.FieldCalculated", util_1.Util.extend(props, properties));
	    };
	    /**
	     * Adds a new SP.FieldDateTime to the collection
	     *
	     * @param title The field title
	     * @param displayFormat The format of the date and time that is displayed in the field.
	     * @param calendarType Specifies the calendar type of the field.
	     * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
	     */
	    Fields.prototype.addDateTime = function (title, displayFormat, calendarType, friendlyDisplayFormat, properties) {
	        if (displayFormat === void 0) { displayFormat = types_1.DateTimeFieldFormatType.DateOnly; }
	        if (calendarType === void 0) { calendarType = types_1.CalendarType.Gregorian; }
	        if (friendlyDisplayFormat === void 0) { friendlyDisplayFormat = 0; }
	        var props = {
	            DateTimeCalendarType: calendarType,
	            DisplayFormat: displayFormat,
	            FieldTypeKind: 4,
	            FriendlyDisplayFormat: friendlyDisplayFormat,
	        };
	        return this.add(title, "SP.FieldDateTime", util_1.Util.extend(props, properties));
	    };
	    /**
	     * Adds a new SP.FieldNumber to the collection
	     *
	     * @param title The field title
	     * @param minValue The field's minimum value
	     * @param maxValue The field's maximum value
	     * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
	     */
	    Fields.prototype.addNumber = function (title, minValue, maxValue, properties) {
	        var props = { FieldTypeKind: 9 };
	        if (typeof minValue !== "undefined") {
	            props = util_1.Util.extend({ MinimumValue: minValue }, props);
	        }
	        if (typeof maxValue !== "undefined") {
	            props = util_1.Util.extend({ MaximumValue: maxValue }, props);
	        }
	        return this.add(title, "SP.FieldNumber", util_1.Util.extend(props, properties));
	    };
	    /**
	     * Adds a new SP.FieldCurrency to the collection
	     *
	     * @param title The field title
	     * @param minValue The field's minimum value
	     * @param maxValue The field's maximum value
	     * @param currencyLocalId Specifies the language code identifier (LCID) used to format the value of the field
	     * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
	     */
	    Fields.prototype.addCurrency = function (title, minValue, maxValue, currencyLocalId, properties) {
	        if (currencyLocalId === void 0) { currencyLocalId = 1033; }
	        var props = {
	            CurrencyLocaleId: currencyLocalId,
	            FieldTypeKind: 10,
	        };
	        if (typeof minValue !== "undefined") {
	            props = util_1.Util.extend({ MinimumValue: minValue }, props);
	        }
	        if (typeof maxValue !== "undefined") {
	            props = util_1.Util.extend({ MaximumValue: maxValue }, props);
	        }
	        return this.add(title, "SP.FieldCurrency", util_1.Util.extend(props, properties));
	    };
	    /**
	     * Adds a new SP.FieldMultiLineText to the collection
	     *
	     * @param title The field title
	     * @param numberOfLines Specifies the number of lines of text to display for the field.
	     * @param richText Specifies whether the field supports rich formatting.
	     * @param restrictedMode Specifies whether the field supports a subset of rich formatting.
	     * @param appendOnly Specifies whether all changes to the value of the field are displayed in list forms.
	     * @param allowHyperlink Specifies whether a hyperlink is allowed as a value of the field.
	     * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
	     *
	     */
	    Fields.prototype.addMultilineText = function (title, numberOfLines, richText, restrictedMode, appendOnly, allowHyperlink, properties) {
	        if (numberOfLines === void 0) { numberOfLines = 6; }
	        if (richText === void 0) { richText = true; }
	        if (restrictedMode === void 0) { restrictedMode = false; }
	        if (appendOnly === void 0) { appendOnly = false; }
	        if (allowHyperlink === void 0) { allowHyperlink = true; }
	        var props = {
	            AllowHyperlink: allowHyperlink,
	            AppendOnly: appendOnly,
	            FieldTypeKind: 3,
	            NumberOfLines: numberOfLines,
	            RestrictedMode: restrictedMode,
	            RichText: richText,
	        };
	        return this.add(title, "SP.FieldMultiLineText", util_1.Util.extend(props, properties));
	    };
	    /**
	     * Adds a new SP.FieldUrl to the collection
	     *
	     * @param title The field title
	     */
	    Fields.prototype.addUrl = function (title, displayFormat, properties) {
	        if (displayFormat === void 0) { displayFormat = types_1.UrlFieldFormatType.Hyperlink; }
	        var props = {
	            DisplayFormat: displayFormat,
	            FieldTypeKind: 11,
	        };
	        return this.add(title, "SP.FieldUrl", util_1.Util.extend(props, properties));
	    };
	    return Fields;
	}(queryable_1.QueryableCollection));
	exports.Fields = Fields;
	/**
	 * Describes a single of Field instance
	 *
	 */
	var Field = (function (_super) {
	    __extends(Field, _super);
	    function Field() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    /**
	     * Updates this field intance with the supplied properties
	     *
	     * @param properties A plain object hash of values to update for the list
	     * @param fieldType The type value, required to update child field type properties
	     */
	    Field.prototype.update = function (properties, fieldType) {
	        var _this = this;
	        if (fieldType === void 0) { fieldType = "SP.Field"; }
	        var postBody = JSON.stringify(util_1.Util.extend({
	            "__metadata": { "type": fieldType },
	        }, properties));
	        return this.post({
	            body: postBody,
	            headers: {
	                "X-HTTP-Method": "MERGE",
	            },
	        }).then(function (data) {
	            return {
	                data: data,
	                field: _this,
	            };
	        });
	    };
	    /**
	     * Delete this fields
	     *
	     */
	    Field.prototype.delete = function () {
	        return this.post({
	            headers: {
	                "X-HTTP-Method": "DELETE",
	            },
	        });
	    };
	    /**
	     * Sets the value of the ShowInDisplayForm property for this field.
	     */
	    Field.prototype.setShowInDisplayForm = function (show) {
	        return this.clone(Field, "setshowindisplayform(" + show + ")", true).post();
	    };
	    /**
	     * Sets the value of the ShowInEditForm property for this field.
	     */
	    Field.prototype.setShowInEditForm = function (show) {
	        return this.clone(Field, "setshowineditform(" + show + ")", true).post();
	    };
	    /**
	     * Sets the value of the ShowInNewForm property for this field.
	     */
	    Field.prototype.setShowInNewForm = function (show) {
	        return this.clone(Field, "setshowinnewform(" + show + ")", true).post();
	    };
	    return Field;
	}(queryable_1.QueryableInstance));
	exports.Field = Field;


/***/ },
/* 47 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	var __extends = (this && this.__extends) || (function () {
	    var extendStatics = Object.setPrototypeOf ||
	        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
	        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
	    return function (d, b) {
	        extendStatics(d, b);
	        function __() { this.constructor = d; }
	        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
	    };
	})();
	Object.defineProperty(exports, "__esModule", { value: true });
	var queryable_1 = __webpack_require__(22);
	/**
	 * Describes a collection of Field objects
	 *
	 */
	var Forms = (function (_super) {
	    __extends(Forms, _super);
	    /**
	     * Creates a new instance of the Fields class
	     *
	     * @param baseUrl The url or Queryable which forms the parent of this fields collection
	     */
	    function Forms(baseUrl, path) {
	        if (path === void 0) { path = "forms"; }
	        return _super.call(this, baseUrl, path) || this;
	    }
	    /**
	     * Gets a form by id
	     *
	     * @param id The guid id of the item to retrieve
	     */
	    Forms.prototype.getById = function (id) {
	        var i = new Form(this);
	        i.concat("('" + id + "')");
	        return i;
	    };
	    return Forms;
	}(queryable_1.QueryableCollection));
	exports.Forms = Forms;
	/**
	 * Describes a single of Form instance
	 *
	 */
	var Form = (function (_super) {
	    __extends(Form, _super);
	    function Form() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    return Form;
	}(queryable_1.QueryableInstance));
	exports.Form = Form;


/***/ },
/* 48 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	var __extends = (this && this.__extends) || (function () {
	    var extendStatics = Object.setPrototypeOf ||
	        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
	        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
	    return function (d, b) {
	        extendStatics(d, b);
	        function __() { this.constructor = d; }
	        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
	    };
	})();
	Object.defineProperty(exports, "__esModule", { value: true });
	var queryable_1 = __webpack_require__(22);
	/**
	 * Describes a collection of webhook subscriptions
	 *
	 */
	var Subscriptions = (function (_super) {
	    __extends(Subscriptions, _super);
	    /**
	     * Creates a new instance of the Subscriptions class
	     *
	     * @param baseUrl - The url or Queryable which forms the parent of this webhook subscriptions collection
	     */
	    function Subscriptions(baseUrl, path) {
	        if (path === void 0) { path = "subscriptions"; }
	        return _super.call(this, baseUrl, path) || this;
	    }
	    /**
	     * Returns all the webhook subscriptions or the specified webhook subscription
	     *
	     */
	    Subscriptions.prototype.getById = function (subscriptionId) {
	        var subscription = new Subscription(this);
	        subscription.concat("('" + subscriptionId + "')");
	        return subscription;
	    };
	    /**
	     * Create a new webhook subscription
	     *
	     */
	    Subscriptions.prototype.add = function (notificationUrl, expirationDate, clientState) {
	        var _this = this;
	        var postBody = JSON.stringify({
	            "clientState": clientState || "pnp-js-core-subscription",
	            "expirationDateTime": expirationDate,
	            "notificationUrl": notificationUrl,
	            "resource": this.toUrl(),
	        });
	        return this.post({ body: postBody, headers: { "Content-Type": "application/json" } }).then(function (result) {
	            return { data: result, subscription: _this.getById(result.id) };
	        });
	    };
	    return Subscriptions;
	}(queryable_1.QueryableCollection));
	exports.Subscriptions = Subscriptions;
	/**
	 * Describes a single webhook subscription instance
	 *
	 */
	var Subscription = (function (_super) {
	    __extends(Subscription, _super);
	    function Subscription() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    /**
	     * Update a webhook subscription
	     *
	     */
	    Subscription.prototype.update = function (expirationDate) {
	        var _this = this;
	        var postBody = JSON.stringify({
	            "expirationDateTime": expirationDate,
	        });
	        return this.patch({ body: postBody, headers: { "Content-Type": "application/json" } }).then(function (data) {
	            return { data: data, subscription: _this };
	        });
	    };
	    /**
	     * Remove a webhook subscription
	     *
	     */
	    Subscription.prototype.delete = function () {
	        return _super.prototype.delete.call(this);
	    };
	    return Subscription;
	}(queryable_1.QueryableInstance));
	exports.Subscription = Subscription;


/***/ },
/* 49 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	var __extends = (this && this.__extends) || (function () {
	    var extendStatics = Object.setPrototypeOf ||
	        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
	        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
	    return function (d, b) {
	        extendStatics(d, b);
	        function __() { this.constructor = d; }
	        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
	    };
	})();
	Object.defineProperty(exports, "__esModule", { value: true });
	var queryable_1 = __webpack_require__(22);
	var util_1 = __webpack_require__(12);
	var UserCustomActions = (function (_super) {
	    __extends(UserCustomActions, _super);
	    function UserCustomActions(baseUrl, path) {
	        if (path === void 0) { path = "usercustomactions"; }
	        return _super.call(this, baseUrl, path) || this;
	    }
	    /**
	     * Returns the custom action with the specified identifier.
	     *
	     * @param id The GUID ID of the user custom action to get.
	     */
	    UserCustomActions.prototype.getById = function (id) {
	        var uca = new UserCustomAction(this);
	        uca.concat("('" + id + "')");
	        return uca;
	    };
	    /**
	     * Create a custom action
	     *
	     * @param creationInfo The information which defines the new custom action
	     *
	     */
	    UserCustomActions.prototype.add = function (properties) {
	        var _this = this;
	        var postBody = JSON.stringify(util_1.Util.extend({ __metadata: { "type": "SP.UserCustomAction" } }, properties));
	        return this.post({ body: postBody }).then(function (data) {
	            return {
	                action: _this.getById(data.Id),
	                data: data,
	            };
	        });
	    };
	    /**
	     * Deletes all custom actions in the collection.
	     *
	     */
	    UserCustomActions.prototype.clear = function () {
	        return this.clone(UserCustomActions, "clear", true).post();
	    };
	    return UserCustomActions;
	}(queryable_1.QueryableCollection));
	exports.UserCustomActions = UserCustomActions;
	var UserCustomAction = (function (_super) {
	    __extends(UserCustomAction, _super);
	    function UserCustomAction() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    UserCustomAction.prototype.update = function (properties) {
	        var _this = this;
	        var postBody = JSON.stringify(util_1.Util.extend({
	            "__metadata": { "type": "SP.UserCustomAction" },
	        }, properties));
	        return this.post({
	            body: postBody,
	            headers: {
	                "X-HTTP-Method": "MERGE",
	            },
	        }).then(function (data) {
	            return {
	                action: _this,
	                data: data,
	            };
	        });
	    };
	    /**
	    * Remove a custom action
	    *
	    */
	    UserCustomAction.prototype.delete = function () {
	        return _super.prototype.delete.call(this);
	    };
	    return UserCustomAction;
	}(queryable_1.QueryableInstance));
	exports.UserCustomAction = UserCustomAction;


/***/ },
/* 50 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	var __extends = (this && this.__extends) || (function () {
	    var extendStatics = Object.setPrototypeOf ||
	        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
	        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
	    return function (d, b) {
	        extendStatics(d, b);
	        function __() { this.constructor = d; }
	        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
	    };
	})();
	Object.defineProperty(exports, "__esModule", { value: true });
	var util_1 = __webpack_require__(12);
	var queryable_1 = __webpack_require__(22);
	/**
	 * Represents a collection of navigation nodes
	 *
	 */
	var NavigationNodes = (function (_super) {
	    __extends(NavigationNodes, _super);
	    function NavigationNodes() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    /**
	     * Gets a navigation node by id
	     *
	     * @param id The id of the node
	     */
	    NavigationNodes.prototype.getById = function (id) {
	        var node = new NavigationNode(this);
	        node.concat("(" + id + ")");
	        return node;
	    };
	    /**
	     * Adds a new node to the collection
	     *
	     * @param title Display name of the node
	     * @param url The url of the node
	     * @param visible If true the node is visible, otherwise it is hidden (default: true)
	     */
	    NavigationNodes.prototype.add = function (title, url, visible) {
	        var _this = this;
	        if (visible === void 0) { visible = true; }
	        var postBody = JSON.stringify({
	            IsVisible: visible,
	            Title: title,
	            Url: url,
	            "__metadata": { "type": "SP.NavigationNode" },
	        });
	        return this.clone(NavigationNodes, null, true).post({ body: postBody }).then(function (data) {
	            return {
	                data: data,
	                node: _this.getById(data.Id),
	            };
	        });
	    };
	    /**
	     * Moves a node to be after another node in the navigation
	     *
	     * @param nodeId Id of the node to move
	     * @param previousNodeId Id of the node after which we move the node specified by nodeId
	     */
	    NavigationNodes.prototype.moveAfter = function (nodeId, previousNodeId) {
	        var postBody = JSON.stringify({
	            nodeId: nodeId,
	            previousNodeId: previousNodeId,
	        });
	        return this.clone(NavigationNodes, "MoveAfter", true).post({ body: postBody });
	    };
	    return NavigationNodes;
	}(queryable_1.QueryableCollection));
	exports.NavigationNodes = NavigationNodes;
	var NavigationNode = (function (_super) {
	    __extends(NavigationNode, _super);
	    function NavigationNode() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    Object.defineProperty(NavigationNode.prototype, "children", {
	        /**
	         * Represents the child nodes of this node
	         */
	        get: function () {
	            return new NavigationNodes(this, "Children");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    /**
	     * Updates this node based on the supplied properties
	     *
	     * @param properties The hash of key/value pairs to update
	     */
	    NavigationNode.prototype.update = function (properties) {
	        var _this = this;
	        var postBody = JSON.stringify(util_1.Util.extend({
	            "__metadata": { "type": "SP.NavigationNode" },
	        }, properties));
	        return this.post({
	            body: postBody,
	            headers: {
	                "X-HTTP-Method": "MERGE",
	            },
	        }).then(function (data) {
	            return {
	                data: data,
	                node: _this,
	            };
	        });
	    };
	    /**
	     * Deletes this node and any child nodes
	     */
	    NavigationNode.prototype.delete = function () {
	        return _super.prototype.delete.call(this);
	    };
	    return NavigationNode;
	}(queryable_1.QueryableInstance));
	exports.NavigationNode = NavigationNode;
	/**
	 * Exposes the navigation components
	 *
	 */
	var Navigation = (function (_super) {
	    __extends(Navigation, _super);
	    /**
	     * Creates a new instance of the Lists class
	     *
	     * @param baseUrl The url or Queryable which forms the parent of this fields collection
	     */
	    function Navigation(baseUrl, path) {
	        if (path === void 0) { path = "navigation"; }
	        return _super.call(this, baseUrl, path) || this;
	    }
	    Object.defineProperty(Navigation.prototype, "quicklaunch", {
	        /**
	         * Gets the quicklaunch navigation for the current context
	         *
	         */
	        get: function () {
	            return new NavigationNodes(this, "quicklaunch");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(Navigation.prototype, "topNavigationBar", {
	        /**
	         * Gets the top bar navigation navigation for the current context
	         *
	         */
	        get: function () {
	            return new NavigationNodes(this, "topnavigationbar");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    return Navigation;
	}(queryable_1.Queryable));
	exports.Navigation = Navigation;


/***/ },
/* 51 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	var __extends = (this && this.__extends) || (function () {
	    var extendStatics = Object.setPrototypeOf ||
	        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
	        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
	    return function (d, b) {
	        extendStatics(d, b);
	        function __() { this.constructor = d; }
	        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
	    };
	})();
	Object.defineProperty(exports, "__esModule", { value: true });
	var queryable_1 = __webpack_require__(22);
	/**
	 * Describes a collection of List objects
	 *
	 */
	var Features = (function (_super) {
	    __extends(Features, _super);
	    /**
	     * Creates a new instance of the Lists class
	     *
	     * @param baseUrl The url or Queryable which forms the parent of this fields collection
	     */
	    function Features(baseUrl, path) {
	        if (path === void 0) { path = "features"; }
	        return _super.call(this, baseUrl, path) || this;
	    }
	    /**
	     * Gets a list from the collection by guid id
	     *
	     * @param id The Id of the feature (GUID)
	     */
	    Features.prototype.getById = function (id) {
	        var feature = new Feature(this);
	        feature.concat("('" + id + "')");
	        return feature;
	    };
	    /**
	     * Adds a new list to the collection
	     *
	     * @param id The Id of the feature (GUID)
	     * @param force If true the feature activation will be forced
	     */
	    Features.prototype.add = function (id, force) {
	        var _this = this;
	        if (force === void 0) { force = false; }
	        return this.clone(Features, "add", true).post({
	            body: JSON.stringify({
	                featdefScope: 0,
	                featureId: id,
	                force: force,
	            }),
	        }).then(function (data) {
	            return {
	                data: data,
	                feature: _this.getById(id),
	            };
	        });
	    };
	    /**
	     * Removes (deactivates) a feature from the collection
	     *
	     * @param id The Id of the feature (GUID)
	     * @param force If true the feature deactivation will be forced
	     */
	    Features.prototype.remove = function (id, force) {
	        if (force === void 0) { force = false; }
	        return this.clone(Features, "remove", true).post({
	            body: JSON.stringify({
	                featureId: id,
	                force: force,
	            }),
	        });
	    };
	    return Features;
	}(queryable_1.QueryableCollection));
	exports.Features = Features;
	var Feature = (function (_super) {
	    __extends(Feature, _super);
	    function Feature() {
	        return _super !== null && _super.apply(this, arguments) || this;
	    }
	    /**
	     * Removes (deactivates) a feature from the collection
	     *
	     * @param force If true the feature deactivation will be forced
	     */
	    Feature.prototype.deactivate = function (force) {
	        var _this = this;
	        if (force === void 0) { force = false; }
	        var removeDependency = this.addBatchDependency();
	        var idGet = new Feature(this).select("DefinitionId");
	        return idGet.getAs().then(function (feature) {
	            var promise = _this.getParent(Features, _this.parentUrl, "", _this.batch).remove(feature.DefinitionId, force);
	            removeDependency();
	            return promise;
	        });
	    };
	    return Feature;
	}(queryable_1.QueryableInstance));
	exports.Feature = Feature;


/***/ },
/* 52 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	var __extends = (this && this.__extends) || (function () {
	    var extendStatics = Object.setPrototypeOf ||
	        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
	        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
	    return function (d, b) {
	        extendStatics(d, b);
	        function __() { this.constructor = d; }
	        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
	    };
	})();
	Object.defineProperty(exports, "__esModule", { value: true });
	var queryable_1 = __webpack_require__(22);
	var RelatedItemManagerImpl = (function (_super) {
	    __extends(RelatedItemManagerImpl, _super);
	    function RelatedItemManagerImpl(baseUrl, path) {
	        if (path === void 0) { path = "_api/SP.RelatedItemManager"; }
	        return _super.call(this, baseUrl, path) || this;
	    }
	    RelatedItemManagerImpl.FromUrl = function (url) {
	        if (url === null) {
	            return new RelatedItemManagerImpl("");
	        }
	        var index = url.indexOf("_api/");
	        if (index > -1) {
	            return new RelatedItemManagerImpl(url.substr(0, index));
	        }
	        return new RelatedItemManagerImpl(url);
	    };
	    RelatedItemManagerImpl.prototype.getRelatedItems = function (sourceListName, sourceItemId) {
	        var query = this.clone(RelatedItemManagerImpl, null, true);
	        query.concat(".GetRelatedItems");
	        return query.post({
	            body: JSON.stringify({
	                SourceItemID: sourceItemId,
	                SourceListName: sourceListName,
	            }),
	        });
	    };
	    RelatedItemManagerImpl.prototype.getPageOneRelatedItems = function (sourceListName, sourceItemId) {
	        var query = this.clone(RelatedItemManagerImpl, null, true);
	        query.concat(".GetPageOneRelatedItems");
	        return query.post({
	            body: JSON.stringify({
	                SourceItemID: sourceItemId,
	                SourceListName: sourceListName,
	            }),
	        });
	    };
	    RelatedItemManagerImpl.prototype.addSingleLink = function (sourceListName, sourceItemId, sourceWebUrl, targetListName, targetItemID, targetWebUrl, tryAddReverseLink) {
	        if (tryAddReverseLink === void 0) { tryAddReverseLink = false; }
	        var query = this.clone(RelatedItemManagerImpl, null, true);
	        query.concat(".AddSingleLink");
	        return query.post({
	            body: JSON.stringify({
	                SourceItemID: sourceItemId,
	                SourceListName: sourceListName,
	                SourceWebUrl: sourceWebUrl,
	                TargetItemID: targetItemID,
	                TargetListName: targetListName,
	                TargetWebUrl: targetWebUrl,
	                TryAddReverseLink: tryAddReverseLink,
	            }),
	        });
	    };
	    /**
	     * Adds a related item link from an item specified by list name and item id, to an item specified by url
	     *
	     * @param sourceListName The source list name or list id
	     * @param sourceItemId The source item id
	     * @param targetItemUrl The target item url
	     * @param tryAddReverseLink If set to true try to add the reverse link (will not return error if it fails)
	     */
	    RelatedItemManagerImpl.prototype.addSingleLinkToUrl = function (sourceListName, sourceItemId, targetItemUrl, tryAddReverseLink) {
	        if (tryAddReverseLink === void 0) { tryAddReverseLink = false; }
	        var query = this.clone(RelatedItemManagerImpl, null, true);
	        query.concat(".AddSingleLinkToUrl");
	        return query.post({
	            body: JSON.stringify({
	                SourceItemID: sourceItemId,
	                SourceListName: sourceListName,
	                TargetItemUrl: targetItemUrl,
	                TryAddReverseLink: tryAddReverseLink,
	            }),
	        });
	    };
	    /**
	     * Adds a related item link from an item specified by url, to an item specified by list name and item id
	     *
	     * @param sourceItemUrl The source item url
	     * @param targetListName The target list name or list id
	     * @param targetItemId The target item id
	     * @param tryAddReverseLink If set to true try to add the reverse link (will not return error if it fails)
	     */
	    RelatedItemManagerImpl.prototype.addSingleLinkFromUrl = function (sourceItemUrl, targetListName, targetItemId, tryAddReverseLink) {
	        if (tryAddReverseLink === void 0) { tryAddReverseLink = false; }
	        var query = this.clone(RelatedItemManagerImpl, null, true);
	        query.concat(".AddSingleLinkFromUrl");
	        return query.post({
	            body: JSON.stringify({
	                SourceItemUrl: sourceItemUrl,
	                TargetItemID: targetItemId,
	                TargetListName: targetListName,
	                TryAddReverseLink: tryAddReverseLink,
	            }),
	        });
	    };
	    RelatedItemManagerImpl.prototype.deleteSingleLink = function (sourceListName, sourceItemId, sourceWebUrl, targetListName, targetItemId, targetWebUrl, tryDeleteReverseLink) {
	        if (tryDeleteReverseLink === void 0) { tryDeleteReverseLink = false; }
	        var query = this.clone(RelatedItemManagerImpl, null, true);
	        query.concat(".DeleteSingleLink");
	        return query.post({
	            body: JSON.stringify({
	                SourceItemID: sourceItemId,
	                SourceListName: sourceListName,
	                SourceWebUrl: sourceWebUrl,
	                TargetItemID: targetItemId,
	                TargetListName: targetListName,
	                TargetWebUrl: targetWebUrl,
	                TryDeleteReverseLink: tryDeleteReverseLink,
	            }),
	        });
	    };
	    return RelatedItemManagerImpl;
	}(queryable_1.Queryable));
	exports.RelatedItemManagerImpl = RelatedItemManagerImpl;


/***/ },
/* 53 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	var __extends = (this && this.__extends) || (function () {
	    var extendStatics = Object.setPrototypeOf ||
	        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
	        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
	    return function (d, b) {
	        extendStatics(d, b);
	        function __() { this.constructor = d; }
	        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
	    };
	})();
	Object.defineProperty(exports, "__esModule", { value: true });
	var queryable_1 = __webpack_require__(22);
	var files_1 = __webpack_require__(54);
	var odata_1 = __webpack_require__(23);
	var UserProfileQuery = (function (_super) {
	    __extends(UserProfileQuery, _super);
	    function UserProfileQuery(baseUrl, path) {
	        if (path === void 0) { path = "_api/sp.userprofiles.peoplemanager"; }
	        var _this = _super.call(this, baseUrl, path) || this;
	        _this.profileLoader = new ProfileLoader(baseUrl);
	        return _this;
	    }
	    Object.defineProperty(UserProfileQuery.prototype, "editProfileLink", {
	        /**
	         * The URL of the edit profile page for the current user.
	         */
	        get: function () {
	            return this.clone(UserProfileQuery, "EditProfileLink").getAs(odata_1.ODataValue());
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(UserProfileQuery.prototype, "isMyPeopleListPublic", {
	        /**
	         * A Boolean value that indicates whether the current user's People I'm Following list is public.
	         */
	        get: function () {
	            return this.clone(UserProfileQuery, "IsMyPeopleListPublic").getAs(odata_1.ODataValue());
	        },
	        enumerable: true,
	        configurable: true
	    });
	    /**
	     * A Boolean value that indicates whether the current user's People I'm Following list is public.
	     *
	     * @param loginName The account name of the user
	     */
	    UserProfileQuery.prototype.amIFollowedBy = function (loginName) {
	        var q = this.clone(UserProfileQuery, "amifollowedby(@v)", true);
	        q.query.add("@v", "'" + encodeURIComponent(loginName) + "'");
	        return q.get();
	    };
	    /**
	     * Checks whether the current user is following the specified user.
	     *
	     * @param loginName The account name of the user
	     */
	    UserProfileQuery.prototype.amIFollowing = function (loginName) {
	        var q = this.clone(UserProfileQuery, "amifollowing(@v)", true);
	        q.query.add("@v", "'" + encodeURIComponent(loginName) + "'");
	        return q.get();
	    };
	    /**
	     * Gets tags that the user is following.
	     *
	     * @param maxCount The maximum number of tags to get.
	     */
	    UserProfileQuery.prototype.getFollowedTags = function (maxCount) {
	        if (maxCount === void 0) { maxCount = 20; }
	        return this.clone(UserProfileQuery, "getfollowedtags(" + maxCount + ")", true).get();
	    };
	    /**
	     * Gets the people who are following the specified user.
	     *
	     * @param loginName The account name of the user.
	     */
	    UserProfileQuery.prototype.getFollowersFor = function (loginName) {
	        var q = this.clone(UserProfileQuery, "getfollowersfor(@v)", true);
	        q.query.add("@v", "'" + encodeURIComponent(loginName) + "'");
	        return q.get();
	    };
	    Object.defineProperty(UserProfileQuery.prototype, "myFollowers", {
	        /**
	         * Gets the people who are following the current user.
	         *
	         */
	        get: function () {
	            return new queryable_1.QueryableCollection(this, "getmyfollowers");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(UserProfileQuery.prototype, "myProperties", {
	        /**
	         * Gets user properties for the current user.
	         *
	         */
	        get: function () {
	            return new UserProfileQuery(this, "getmyproperties");
	        },
	        enumerable: true,
	        configurable: true
	    });
	    /**
	     * Gets the people who the specified user is following.
	     *
	     * @param loginName The account name of the user.
	     */
	    UserProfileQuery.prototype.getPeopleFollowedBy = function (loginName) {
	        var q = this.clone(UserProfileQuery, "getpeoplefollowedby(@v)", true);
	        q.query.add("@v", "'" + encodeURIComponent(loginName) + "'");
	        return q.get();
	    };
	    /**
	     * Gets user properties for the specified user.
	     *
	     * @param loginName The account name of the user.
	     */
	    UserProfileQuery.prototype.getPropertiesFor = function (loginName) {
	        var q = this.clone(UserProfileQuery, "getpropertiesfor(@v)", true);
	        q.query.add("@v", "'" + encodeURIComponent(loginName) + "'");
	        return q.get();
	    };
	    Object.defineProperty(UserProfileQuery.prototype, "trendingTags", {
	        /**
	         * Gets the most popular tags.
	         *
	         */
	        get: function () {
	            var q = this.clone(UserProfileQuery, null, true);
	            q.concat(".gettrendingtags");
	            return q.get();
	        },
	        enumerable: true,
	        configurable: true
	    });
	    /**
	     * Gets the specified user profile property for the specified user.
	     *
	     * @param loginName The account name of the user.
	     * @param propertyName The case-sensitive name of the property to get.
	     */
	    UserProfileQuery.prototype.getUserProfilePropertyFor = function (loginName, propertyName) {
	        var q = this.clone(UserProfileQuery, "getuserprofilepropertyfor(accountname=@v, propertyname='" + propertyName + "')", true);
	        q.query.add("@v", "'" + encodeURIComponent(loginName) + "'");
	        return q.get();
	    };
	    /**
	     * Removes the specified user from the user's list of suggested people to follow.
	     *
	     * @param loginName The account name of the user.
	     */
	    UserProfileQuery.prototype.hideSuggestion = function (loginName) {
	        var q = this.clone(UserProfileQuery, "hidesuggestion(@v)", true);
	        q.query.add("@v", "'" + encodeURIComponent(loginName) + "'");
	        return q.post();
	    };
	    /**
	     * Checks whether the first user is following the second user.
	     *
	     * @param follower The account name of the user who might be following followee.
	     * @param followee The account name of the user who might be followed.
	     */
	    UserProfileQuery.prototype.isFollowing = function (follower, followee) {
	        var q = this.clone(UserProfileQuery, null, true);
	        q.concat(".isfollowing(possiblefolloweraccountname=@v, possiblefolloweeaccountname=@y)");
	        q.query.add("@v", "'" + encodeURIComponent(follower) + "'");
	        q.query.add("@y", "'" + encodeURIComponent(followee) + "'");
	        return q.get();
	    };
	    /**
	     * Uploads and sets the user profile picture. Not supported for batching.
	     *
	     * @param profilePicSource Blob data representing the user's picture
	     */
	    UserProfileQuery.prototype.setMyProfilePic = function (profilePicSource) {
	        var _this = this;
	        return new Promise(function (resolve, reject) {
	            files_1.readBlobAsArrayBuffer(profilePicSource).then(function (buffer) {
	                var request = new UserProfileQuery(_this, "setmyprofilepicture");
	                request.post({
	                    body: String.fromCharCode.apply(null, new Uint16Array(buffer)),
	                }).then(function (_) { return resolve(); });
	            }).catch(function (e) { return reject(e); });
	        });
	    };
	    /**
	     * Provisions one or more users' personal sites. (My Site administrator on SharePoint Online only)
	     *
	     * @param emails The email addresses of the users to provision sites for
	     */
	    UserProfileQuery.prototype.createPersonalSiteEnqueueBulk = function () {
	        var emails = [];
	        for (var _i = 0; _i < arguments.length; _i++) {
	            emails[_i] = arguments[_i];
	        }
	        return this.profileLoader.createPersonalSiteEnqueueBulk(emails);
	    };
	    Object.defineProperty(UserProfileQuery.prototype, "ownerUserProfile", {
	        /**
	         * Gets the user profile of the site owner.
	         *
	         */
	        get: function () {
	            return this.profileLoader.ownerUserProfile;
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(UserProfileQuery.prototype, "userProfile", {
	        /**
	         * Gets the user profile that corresponds to the current user.
	         */
	        get: function () {
	            return this.profileLoader.userProfile;
	        },
	        enumerable: true,
	        configurable: true
	    });
	    /**
	     * Enqueues creating a personal site for this user, which can be used to share documents, web pages, and other files.
	     *
	     * @param interactiveRequest true if interactively (web) initiated request, or false if non-interactively (client) initiated request
	     */
	    UserProfileQuery.prototype.createPersonalSite = function (interactiveRequest) {
	        if (interactiveRequest === void 0) { interactiveRequest = false; }
	        return this.profileLoader.createPersonalSite(interactiveRequest);
	    };
	    /**
	     * Sets the privacy settings for this profile.
	     *
	     * @param share true to make all social data public; false to make all social data private.
	     */
	    UserProfileQuery.prototype.shareAllSocialData = function (share) {
	        return this.profileLoader.shareAllSocialData(share);
	    };
	    return UserProfileQuery;
	}(queryable_1.QueryableInstance));
	exports.UserProfileQuery = UserProfileQuery;
	var ProfileLoader = (function (_super) {
	    __extends(ProfileLoader, _super);
	    function ProfileLoader(baseUrl, path) {
	        if (path === void 0) { path = "_api/sp.userprofiles.profileloader.getprofileloader"; }
	        return _super.call(this, baseUrl, path) || this;
	    }
	    /**
	     * Provisions one or more users' personal sites. (My Site administrator on SharePoint Online only)
	     *
	     * @param emails The email addresses of the users to provision sites for
	     */
	    ProfileLoader.prototype.createPersonalSiteEnqueueBulk = function (emails) {
	        return this.clone(ProfileLoader, "createpersonalsiteenqueuebulk").post({
	            body: JSON.stringify({ "emailIDs": emails }),
	        });
	    };
	    Object.defineProperty(ProfileLoader.prototype, "ownerUserProfile", {
	        /**
	         * Gets the user profile of the site owner.
	         *
	         */
	        get: function () {
	            var q = this.getParent(ProfileLoader, this.parentUrl, "_api/sp.userprofiles.profileloader.getowneruserprofile");
	            if (this.hasBatch) {
	                q = q.inBatch(this.batch);
	            }
	            return q.postAs();
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(ProfileLoader.prototype, "userProfile", {
	        /**
	         * Gets the user profile that corresponds to the current user.
	         *
	         */
	        get: function () {
	            return this.clone(ProfileLoader, "getuserprofile", true).postAs();
	        },
	        enumerable: true,
	        configurable: true
	    });
	    /**
	     * Enqueues creating a personal site for this user, which can be used to share documents, web pages, and other files.
	     *
	     * @param interactiveRequest true if interactively (web) initiated request, or false if non-interactively (client) initiated request
	     */
	    ProfileLoader.prototype.createPersonalSite = function (interactiveRequest) {
	        if (interactiveRequest === void 0) { interactiveRequest = false; }
	        return this.clone(ProfileLoader, "getuserprofile/createpersonalsiteenque(" + interactiveRequest + ")", true).post();
	    };
	    /**
	     * Sets the privacy settings for this profile.
	     *
	     * @param share true to make all social data public; false to make all social data private.
	     */
	    ProfileLoader.prototype.shareAllSocialData = function (share) {
	        return this.clone(ProfileLoader, "getuserprofile/shareallsocialdata(" + share + ")", true).post();
	    };
	    return ProfileLoader;
	}(queryable_1.Queryable));


/***/ },
/* 54 */
/***/ function(module, exports) {

	"use strict";
	Object.defineProperty(exports, "__esModule", { value: true });
	/**
	 * Reads a blob as text
	 *
	 * @param blob The data to read
	 */
	function readBlobAsText(blob) {
	    return readBlobAs(blob, "string");
	}
	exports.readBlobAsText = readBlobAsText;
	/**
	 * Reads a blob into an array buffer
	 *
	 * @param blob The data to read
	 */
	function readBlobAsArrayBuffer(blob) {
	    return readBlobAs(blob, "buffer");
	}
	exports.readBlobAsArrayBuffer = readBlobAsArrayBuffer;
	/**
	 * Generic method to read blob's content
	 *
	 * @param blob The data to read
	 * @param mode The read mode
	 */
	function readBlobAs(blob, mode) {
	    return new Promise(function (resolve, reject) {
	        try {
	            var reader = new FileReader();
	            reader.onload = function (e) {
	                resolve(e.target.result);
	            };
	            switch (mode) {
	                case "string":
	                    reader.readAsText(blob);
	                    break;
	                case "buffer":
	                    reader.readAsArrayBuffer(blob);
	                    break;
	            }
	        }
	        catch (e) {
	            reject(e);
	        }
	    });
	}


/***/ },
/* 55 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	function __export(m) {
	    for (var p in m) if (!exports.hasOwnProperty(p)) exports[p] = m[p];
	}
	Object.defineProperty(exports, "__esModule", { value: true });
	__export(__webpack_require__(56));
	var httpclient_1 = __webpack_require__(24);
	exports.HttpClient = httpclient_1.HttpClient;
	var sprequestexecutorclient_1 = __webpack_require__(57);
	exports.SPRequestExecutorClient = sprequestexecutorclient_1.SPRequestExecutorClient;
	var nodefetchclient_1 = __webpack_require__(58);
	exports.NodeFetchClient = nodefetchclient_1.NodeFetchClient;
	var fetchclient_1 = __webpack_require__(16);
	exports.FetchClient = fetchclient_1.FetchClient;
	__export(__webpack_require__(59));
	var collections_1 = __webpack_require__(18);
	exports.Dictionary = collections_1.Dictionary;
	var util_1 = __webpack_require__(12);
	exports.Util = util_1.Util;
	__export(__webpack_require__(14));
	__export(__webpack_require__(26));


/***/ },
/* 56 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	function __export(m) {
	    for (var p in m) if (!exports.hasOwnProperty(p)) exports[p] = m[p];
	}
	Object.defineProperty(exports, "__esModule", { value: true });
	__export(__webpack_require__(28));
	var files_1 = __webpack_require__(41);
	exports.CheckinType = files_1.CheckinType;
	exports.WebPartsPersonalizationScope = files_1.WebPartsPersonalizationScope;
	exports.MoveOperations = files_1.MoveOperations;
	exports.TemplateFileType = files_1.TemplateFileType;
	var folders_1 = __webpack_require__(40);
	exports.Folder = folders_1.Folder;
	exports.Folders = folders_1.Folders;
	var items_1 = __webpack_require__(33);
	exports.Item = items_1.Item;
	exports.Items = items_1.Items;
	exports.PagedItemCollection = items_1.PagedItemCollection;
	var navigation_1 = __webpack_require__(50);
	exports.NavigationNodes = navigation_1.NavigationNodes;
	exports.NavigationNode = navigation_1.NavigationNode;
	var lists_1 = __webpack_require__(32);
	exports.List = lists_1.List;
	exports.Lists = lists_1.Lists;
	var odata_1 = __webpack_require__(23);
	exports.extractOdataId = odata_1.extractOdataId;
	exports.ODataParserBase = odata_1.ODataParserBase;
	exports.ODataDefaultParser = odata_1.ODataDefaultParser;
	exports.ODataRaw = odata_1.ODataRaw;
	exports.ODataValue = odata_1.ODataValue;
	exports.ODataEntity = odata_1.ODataEntity;
	exports.ODataEntityArray = odata_1.ODataEntityArray;
	exports.TextFileParser = odata_1.TextFileParser;
	exports.BlobFileParser = odata_1.BlobFileParser;
	exports.BufferFileParser = odata_1.BufferFileParser;
	exports.JSONFileParser = odata_1.JSONFileParser;
	var queryable_1 = __webpack_require__(22);
	exports.Queryable = queryable_1.Queryable;
	exports.QueryableInstance = queryable_1.QueryableInstance;
	exports.QueryableCollection = queryable_1.QueryableCollection;
	var roles_1 = __webpack_require__(36);
	exports.RoleDefinitionBindings = roles_1.RoleDefinitionBindings;
	var search_1 = __webpack_require__(21);
	exports.Search = search_1.Search;
	exports.SearchResult = search_1.SearchResult;
	exports.SearchResults = search_1.SearchResults;
	exports.SortDirection = search_1.SortDirection;
	exports.ReorderingRuleMatchType = search_1.ReorderingRuleMatchType;
	exports.QueryPropertyValueType = search_1.QueryPropertyValueType;
	var searchsuggest_1 = __webpack_require__(29);
	exports.SearchSuggest = searchsuggest_1.SearchSuggest;
	exports.SearchSuggestResult = searchsuggest_1.SearchSuggestResult;
	var site_1 = __webpack_require__(30);
	exports.Site = site_1.Site;
	__export(__webpack_require__(39));
	var webs_1 = __webpack_require__(31);
	exports.Web = webs_1.Web;


/***/ },
/* 57 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	Object.defineProperty(exports, "__esModule", { value: true });
	var util_1 = __webpack_require__(12);
	var exceptions_1 = __webpack_require__(26);
	/**
	 * Makes requests using the SP.RequestExecutor library.
	 */
	var SPRequestExecutorClient = (function () {
	    function SPRequestExecutorClient() {
	        /**
	         * Converts a SharePoint REST API response to a fetch API response.
	         */
	        this.convertToResponse = function (spResponse) {
	            var responseHeaders = new Headers();
	            for (var h in spResponse.headers) {
	                if (spResponse.headers[h]) {
	                    responseHeaders.append(h, spResponse.headers[h]);
	                }
	            }
	            // issue #256, Cannot have an empty string body when creating a Response with status 204
	            var body = spResponse.statusCode === 204 ? null : spResponse.body;
	            return new Response(body, {
	                headers: responseHeaders,
	                status: spResponse.statusCode,
	                statusText: spResponse.statusText,
	            });
	        };
	    }
	    /**
	     * Fetches a URL using the SP.RequestExecutor library.
	     */
	    SPRequestExecutorClient.prototype.fetch = function (url, options) {
	        var _this = this;
	        if (typeof SP === "undefined" || typeof SP.RequestExecutor === "undefined") {
	            throw new exceptions_1.SPRequestExecutorUndefinedException();
	        }
	        var addinWebUrl = url.substring(0, url.indexOf("/_api")), executor = new SP.RequestExecutor(addinWebUrl);
	        var headers = {}, iterator, temp;
	        if (options.headers && options.headers instanceof Headers) {
	            iterator = options.headers.entries();
	            temp = iterator.next();
	            while (!temp.done) {
	                headers[temp.value[0]] = temp.value[1];
	                temp = iterator.next();
	            }
	        }
	        else {
	            headers = options.headers;
	        }
	        return new Promise(function (resolve, reject) {
	            var requestOptions = {
	                error: function (error) {
	                    reject(_this.convertToResponse(error));
	                },
	                headers: headers,
	                method: options.method,
	                success: function (response) {
	                    resolve(_this.convertToResponse(response));
	                },
	                url: url,
	            };
	            if (options.body) {
	                requestOptions = util_1.Util.extend(requestOptions, { body: options.body });
	            }
	            else {
	                requestOptions = util_1.Util.extend(requestOptions, { binaryStringRequestBody: true });
	            }
	            executor.executeAsync(requestOptions);
	        });
	    };
	    return SPRequestExecutorClient;
	}());
	exports.SPRequestExecutorClient = SPRequestExecutorClient;


/***/ },
/* 58 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	Object.defineProperty(exports, "__esModule", { value: true });
	var exceptions_1 = __webpack_require__(26);
	/**
	 * This module is substituted for the NodeFetchClient.ts during the packaging process. This helps to reduce the pnp.js file size by
	 * not including all of the node dependencies
	 */
	var NodeFetchClient = (function () {
	    function NodeFetchClient() {
	    }
	    /**
	     * Always throws an error that NodeFetchClient is not supported for use in the browser
	     */
	    NodeFetchClient.prototype.fetch = function () {
	        throw new exceptions_1.NodeFetchClientUnsupportedException();
	    };
	    return NodeFetchClient;
	}());
	exports.NodeFetchClient = NodeFetchClient;


/***/ },
/* 59 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	Object.defineProperty(exports, "__esModule", { value: true });
	var cachingConfigurationProvider_1 = __webpack_require__(60);
	exports.CachingConfigurationProvider = cachingConfigurationProvider_1.default;
	var spListConfigurationProvider_1 = __webpack_require__(61);
	exports.SPListConfigurationProvider = spListConfigurationProvider_1.default;


/***/ },
/* 60 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	Object.defineProperty(exports, "__esModule", { value: true });
	var storage = __webpack_require__(17);
	var exceptions_1 = __webpack_require__(26);
	/**
	 * A caching provider which can wrap other non-caching providers
	 *
	 */
	var CachingConfigurationProvider = (function () {
	    /**
	     * Creates a new caching configuration provider
	     * @constructor
	     * @param {IConfigurationProvider} wrappedProvider Provider which will be used to fetch the configuration
	     * @param {string} cacheKey Key that will be used to store cached items to the cache
	     * @param {IPnPClientStore} cacheStore OPTIONAL storage, which will be used to store cached settings.
	     */
	    function CachingConfigurationProvider(wrappedProvider, cacheKey, cacheStore) {
	        this.wrappedProvider = wrappedProvider;
	        this.store = (cacheStore) ? cacheStore : this.selectPnPCache();
	        this.cacheKey = "_configcache_" + cacheKey;
	    }
	    /**
	     * Gets the wrapped configuration providers
	     *
	     * @return {IConfigurationProvider} Wrapped configuration provider
	     */
	    CachingConfigurationProvider.prototype.getWrappedProvider = function () {
	        return this.wrappedProvider;
	    };
	    /**
	     * Loads the configuration values either from the cache or from the wrapped provider
	     *
	     * @return {Promise<TypedHash<string>>} Promise of loaded configuration values
	     */
	    CachingConfigurationProvider.prototype.getConfiguration = function () {
	        var _this = this;
	        // Cache not available, pass control to  the wrapped provider
	        if ((!this.store) || (!this.store.enabled)) {
	            return this.wrappedProvider.getConfiguration();
	        }
	        // Value is found in cache, return it directly
	        var cachedConfig = this.store.get(this.cacheKey);
	        if (cachedConfig) {
	            return new Promise(function (resolve) {
	                resolve(cachedConfig);
	            });
	        }
	        // Get and cache value from the wrapped provider
	        var providerPromise = this.wrappedProvider.getConfiguration();
	        providerPromise.then(function (providedConfig) {
	            _this.store.put(_this.cacheKey, providedConfig);
	        });
	        return providerPromise;
	    };
	    CachingConfigurationProvider.prototype.selectPnPCache = function () {
	        var pnpCache = new storage.PnPClientStorage();
	        if ((pnpCache.local) && (pnpCache.local.enabled)) {
	            return pnpCache.local;
	        }
	        if ((pnpCache.session) && (pnpCache.session.enabled)) {
	            return pnpCache.session;
	        }
	        throw new exceptions_1.NoCacheAvailableException();
	    };
	    return CachingConfigurationProvider;
	}());
	exports.default = CachingConfigurationProvider;


/***/ },
/* 61 */
/***/ function(module, exports, __webpack_require__) {

	"use strict";
	Object.defineProperty(exports, "__esModule", { value: true });
	var cachingConfigurationProvider_1 = __webpack_require__(60);
	/**
	 * A configuration provider which loads configuration values from a SharePoint list
	 *
	 */
	var SPListConfigurationProvider = (function () {
	    /**
	     * Creates a new SharePoint list based configuration provider
	     * @constructor
	     * @param {string} webUrl Url of the SharePoint site, where the configuration list is located
	     * @param {string} listTitle Title of the SharePoint list, which contains the configuration settings (optional, default = "config")
	     */
	    function SPListConfigurationProvider(sourceWeb, sourceListTitle) {
	        if (sourceListTitle === void 0) { sourceListTitle = "config"; }
	        this.sourceWeb = sourceWeb;
	        this.sourceListTitle = sourceListTitle;
	    }
	    Object.defineProperty(SPListConfigurationProvider.prototype, "web", {
	        /**
	         * Gets the url of the SharePoint site, where the configuration list is located
	         *
	         * @return {string} Url address of the site
	         */
	        get: function () {
	            return this.sourceWeb;
	        },
	        enumerable: true,
	        configurable: true
	    });
	    Object.defineProperty(SPListConfigurationProvider.prototype, "listTitle", {
	        /**
	         * Gets the title of the SharePoint list, which contains the configuration settings
	         *
	         * @return {string} List title
	         */
	        get: function () {
	            return this.sourceListTitle;
	        },
	        enumerable: true,
	        configurable: true
	    });
	    /**
	     * Loads the configuration values from the SharePoint list
	     *
	     * @return {Promise<TypedHash<string>>} Promise of loaded configuration values
	     */
	    SPListConfigurationProvider.prototype.getConfiguration = function () {
	        return this.web.lists.getByTitle(this.listTitle).items.select("Title", "Value")
	            .getAs().then(function (data) {
	            return data.reduce(function (configuration, item) {
	                return Object.defineProperty(configuration, item.Title, {
	                    configurable: false,
	                    enumerable: false,
	                    value: item.Value,
	                    writable: false,
	                });
	            }, {});
	        });
	    };
	    /**
	     * Wraps the current provider in a cache enabled provider
	     *
	     * @return {CachingConfigurationProvider} Caching providers which wraps the current provider
	     */
	    SPListConfigurationProvider.prototype.asCaching = function () {
	        var cacheKey = "splist_" + this.web.toUrl() + "+" + this.listTitle;
	        return new cachingConfigurationProvider_1.default(this, cacheKey);
	    };
	    return SPListConfigurationProvider;
	}());
	exports.default = SPListConfigurationProvider;


/***/ }
/******/ ])});;
//# sourceMappingURL=quick-links.bundle.js.map