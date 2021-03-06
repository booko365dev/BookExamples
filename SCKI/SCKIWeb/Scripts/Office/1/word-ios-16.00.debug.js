/* Word iOS-specific API library */
/* Version: 16.0.9105.3000 */

/* Office.js Version: 16.0.9124.1000 */ 
/*
	Copyright (c) Microsoft Corporation.  All rights reserved.
*/

/*
	Your use of this file is governed by the Microsoft Services Agreement http://go.microsoft.com/fwlink/?LinkId=266419.
*/

/*
* @overview es6-promise - a tiny implementation of Promises/A+.
* @copyright Copyright (c) 2014 Yehuda Katz, Tom Dale, Stefan Penner and contributors (Conversion to ES6 API by Jake Archibald)
* @license   Licensed under MIT license
*            See https://raw.githubusercontent.com/jakearchibald/es6-promise/master/LICENSE
* @version   2.3.0
*/

var __extends=(this && this.__extends) || function (d, b) {
	for (var p in b) if (b.hasOwnProperty(p)) d[p]=b[p];
	function __() { this.constructor=d; }
	d.prototype=b===null ? Object.create(b) : (__.prototype=b.prototype, new __());
};
var OfficeExt;
(function (OfficeExt) {
	var MicrosoftAjaxFactory=(function () {
		function MicrosoftAjaxFactory() {
		}
		MicrosoftAjaxFactory.prototype.isMsAjaxLoaded=function () {
			if (typeof (Sys) !=='undefined' && typeof (Type) !=='undefined' &&
				Sys.StringBuilder && typeof (Sys.StringBuilder)==="function" &&
				Type.registerNamespace && typeof (Type.registerNamespace)==="function" &&
				Type.registerClass && typeof (Type.registerClass)==="function" &&
				typeof (Function._validateParams)==="function" &&
				Sys.Serialization && Sys.Serialization.JavaScriptSerializer && typeof (Sys.Serialization.JavaScriptSerializer.serialize)==="function") {
				return true;
			}
			else {
				return false;
			}
		};
		MicrosoftAjaxFactory.prototype.loadMsAjaxFull=function (callback) {
			var msAjaxCDNPath=(window.location.protocol.toLowerCase()==='https:' ? 'https:' : 'http:')+'//ajax.aspnetcdn.com/ajax/3.5/MicrosoftAjax.js';
			OSF.OUtil.loadScript(msAjaxCDNPath, callback);
		};
		Object.defineProperty(MicrosoftAjaxFactory.prototype, "msAjaxError", {
			get: function () {
				if (this._msAjaxError==null && this.isMsAjaxLoaded()) {
					this._msAjaxError=Error;
				}
				return this._msAjaxError;
			},
			set: function (errorClass) {
				this._msAjaxError=errorClass;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(MicrosoftAjaxFactory.prototype, "msAjaxString", {
			get: function () {
				if (this._msAjaxString==null && this.isMsAjaxLoaded()) {
					this._msAjaxString=String;
				}
				return this._msAjaxString;
			},
			set: function (stringClass) {
				this._msAjaxString=stringClass;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(MicrosoftAjaxFactory.prototype, "msAjaxDebug", {
			get: function () {
				if (this._msAjaxDebug==null && this.isMsAjaxLoaded()) {
					this._msAjaxDebug=Sys.Debug;
				}
				return this._msAjaxDebug;
			},
			set: function (debugClass) {
				this._msAjaxDebug=debugClass;
			},
			enumerable: true,
			configurable: true
		});
		return MicrosoftAjaxFactory;
	})();
	OfficeExt.MicrosoftAjaxFactory=MicrosoftAjaxFactory;
})(OfficeExt || (OfficeExt={}));
var OsfMsAjaxFactory=new OfficeExt.MicrosoftAjaxFactory();
var OSF=OSF || {};
var OfficeExt;
(function (OfficeExt) {
	var SafeStorage=(function () {
		function SafeStorage(_internalStorage) {
			this._internalStorage=_internalStorage;
		}
		SafeStorage.prototype.getItem=function (key) {
			try {
				return this._internalStorage && this._internalStorage.getItem(key);
			}
			catch (e) {
				return null;
			}
		};
		SafeStorage.prototype.setItem=function (key, data) {
			try {
				this._internalStorage && this._internalStorage.setItem(key, data);
			}
			catch (e) {
			}
		};
		SafeStorage.prototype.clear=function () {
			try {
				this._internalStorage && this._internalStorage.clear();
			}
			catch (e) {
			}
		};
		SafeStorage.prototype.removeItem=function (key) {
			try {
				this._internalStorage && this._internalStorage.removeItem(key);
			}
			catch (e) {
			}
		};
		SafeStorage.prototype.getKeysWithPrefix=function (keyPrefix) {
			var keyList=[];
			try {
				var len=this._internalStorage && this._internalStorage.length || 0;
				for (var i=0; i < len; i++) {
					var key=this._internalStorage.key(i);
					if (key.indexOf(keyPrefix)===0) {
						keyList.push(key);
					}
				}
			}
			catch (e) {
			}
			return keyList;
		};
		return SafeStorage;
	})();
	OfficeExt.SafeStorage=SafeStorage;
})(OfficeExt || (OfficeExt={}));
OSF.XdmFieldName={
	ConversationUrl: "ConversationUrl",
	AppId: "AppId"
};
OSF.WindowNameItemKeys={
	BaseFrameName: "baseFrameName",
	HostInfo: "hostInfo",
	XdmInfo: "xdmInfo",
	SerializerVersion: "serializerVersion",
	AppContext: "appContext"
};
OSF.OUtil=(function () {
	var _uniqueId=-1;
	var _xdmInfoKey='&_xdm_Info=';
	var _serializerVersionKey='&_serializer_version=';
	var _xdmSessionKeyPrefix='_xdm_';
	var _serializerVersionKeyPrefix='_serializer_version=';
	var _fragmentSeparator='#';
	var _fragmentInfoDelimiter='&';
	var _classN="class";
	var _loadedScripts={};
	var _defaultScriptLoadingTimeout=30000;
	var _safeSessionStorage=null;
	var _safeLocalStorage=null;
	var _rndentropy=new Date().getTime();
	function _random() {
		var nextrand=0x7fffffff * (Math.random());
		nextrand ^=_rndentropy ^ ((new Date().getMilliseconds()) << Math.floor(Math.random() * (31 - 10)));
		return nextrand.toString(16);
	}
	;
	function _getSessionStorage() {
		if (!_safeSessionStorage) {
			try {
				var sessionStorage=window.sessionStorage;
			}
			catch (ex) {
				sessionStorage=null;
			}
			_safeSessionStorage=new OfficeExt.SafeStorage(sessionStorage);
		}
		return _safeSessionStorage;
	}
	;
	function _reOrderTabbableElements(elements) {
		var bucket0=[];
		var bucketPositive=[];
		var i;
		var len=elements.length;
		var ele;
		for (i=0; i < len; i++) {
			ele=elements[i];
			if (ele.tabIndex) {
				if (ele.tabIndex > 0) {
					bucketPositive.push(ele);
				}
				else if (ele.tabIndex===0) {
					bucket0.push(ele);
				}
			}
			else {
				bucket0.push(ele);
			}
		}
		bucketPositive=bucketPositive.sort(function (left, right) {
			var diff=left.tabIndex - right.tabIndex;
			if (diff===0) {
				diff=bucketPositive.indexOf(left) - bucketPositive.indexOf(right);
			}
			return diff;
		});
		return [].concat(bucketPositive, bucket0);
	}
	;
	return {
		set_entropy: function OSF_OUtil$set_entropy(entropy) {
			if (typeof entropy=="string") {
				for (var i=0; i < entropy.length; i+=4) {
					var temp=0;
					for (var j=0; j < 4 && i+j < entropy.length; j++) {
						temp=(temp << 8)+entropy.charCodeAt(i+j);
					}
					_rndentropy ^=temp;
				}
			}
			else if (typeof entropy=="number") {
				_rndentropy ^=entropy;
			}
			else {
				_rndentropy ^=0x7fffffff * Math.random();
			}
			_rndentropy &=0x7fffffff;
		},
		extend: function OSF_OUtil$extend(child, parent) {
			var F=function () { };
			F.prototype=parent.prototype;
			child.prototype=new F();
			child.prototype.constructor=child;
			child.uber=parent.prototype;
			if (parent.prototype.constructor===Object.prototype.constructor) {
				parent.prototype.constructor=parent;
			}
		},
		setNamespace: function OSF_OUtil$setNamespace(name, parent) {
			if (parent && name && !parent[name]) {
				parent[name]={};
			}
		},
		unsetNamespace: function OSF_OUtil$unsetNamespace(name, parent) {
			if (parent && name && parent[name]) {
				delete parent[name];
			}
		},
		serializeSettings: function OSF_OUtil$serializeSettings(settingsCollection) {
			var ret={};
			for (var key in settingsCollection) {
				var value=settingsCollection[key];
				try {
					if (JSON) {
						value=JSON.stringify(value, function dateReplacer(k, v) {
							return OSF.OUtil.isDate(this[k]) ? OSF.DDA.SettingsManager.DateJSONPrefix+this[k].getTime()+OSF.DDA.SettingsManager.DataJSONSuffix : v;
						});
					}
					else {
						value=Sys.Serialization.JavaScriptSerializer.serialize(value);
					}
					ret[key]=value;
				}
				catch (ex) {
				}
			}
			return ret;
		},
		deserializeSettings: function OSF_OUtil$deserializeSettings(serializedSettings) {
			var ret={};
			serializedSettings=serializedSettings || {};
			for (var key in serializedSettings) {
				var value=serializedSettings[key];
				try {
					if (JSON) {
						value=JSON.parse(value, function dateReviver(k, v) {
							var d;
							if (typeof v==='string' && v && v.length > 6 && v.slice(0, 5)===OSF.DDA.SettingsManager.DateJSONPrefix && v.slice(-1)===OSF.DDA.SettingsManager.DataJSONSuffix) {
								d=new Date(parseInt(v.slice(5, -1)));
								if (d) {
									return d;
								}
							}
							return v;
						});
					}
					else {
						value=Sys.Serialization.JavaScriptSerializer.deserialize(value, true);
					}
					ret[key]=value;
				}
				catch (ex) {
				}
			}
			return ret;
		},
		loadScript: function OSF_OUtil$loadScript(url, callback, timeoutInMs) {
			if (url && callback) {
				var doc=window.document;
				var _loadedScriptEntry=_loadedScripts[url];
				if (!_loadedScriptEntry) {
					var script=doc.createElement("script");
					script.type="text/javascript";
					_loadedScriptEntry={ loaded: false, pendingCallbacks: [callback], timer: null };
					_loadedScripts[url]=_loadedScriptEntry;
					var onLoadCallback=function OSF_OUtil_loadScript$onLoadCallback() {
						if (_loadedScriptEntry.timer !=null) {
							clearTimeout(_loadedScriptEntry.timer);
							delete _loadedScriptEntry.timer;
						}
						_loadedScriptEntry.loaded=true;
						var pendingCallbackCount=_loadedScriptEntry.pendingCallbacks.length;
						for (var i=0; i < pendingCallbackCount; i++) {
							var currentCallback=_loadedScriptEntry.pendingCallbacks.shift();
							currentCallback();
						}
					};
					var onLoadError=function OSF_OUtil_loadScript$onLoadError() {
						delete _loadedScripts[url];
						if (_loadedScriptEntry.timer !=null) {
							clearTimeout(_loadedScriptEntry.timer);
							delete _loadedScriptEntry.timer;
						}
						var pendingCallbackCount=_loadedScriptEntry.pendingCallbacks.length;
						for (var i=0; i < pendingCallbackCount; i++) {
							var currentCallback=_loadedScriptEntry.pendingCallbacks.shift();
							currentCallback();
						}
					};
					if (script.readyState) {
						script.onreadystatechange=function () {
							if (script.readyState=="loaded" || script.readyState=="complete") {
								script.onreadystatechange=null;
								onLoadCallback();
							}
						};
					}
					else {
						script.onload=onLoadCallback;
					}
					script.onerror=onLoadError;
					timeoutInMs=timeoutInMs || _defaultScriptLoadingTimeout;
					_loadedScriptEntry.timer=setTimeout(onLoadError, timeoutInMs);
					script.setAttribute("crossOrigin", "anonymous");
					script.src=url;
					doc.getElementsByTagName("head")[0].appendChild(script);
				}
				else if (_loadedScriptEntry.loaded) {
					callback();
				}
				else {
					_loadedScriptEntry.pendingCallbacks.push(callback);
				}
			}
		},
		loadCSS: function OSF_OUtil$loadCSS(url) {
			if (url) {
				var doc=window.document;
				var link=doc.createElement("link");
				link.type="text/css";
				link.rel="stylesheet";
				link.href=url;
				doc.getElementsByTagName("head")[0].appendChild(link);
			}
		},
		parseEnum: function OSF_OUtil$parseEnum(str, enumObject) {
			var parsed=enumObject[str.trim()];
			if (typeof (parsed)=='undefined') {
				OsfMsAjaxFactory.msAjaxDebug.trace("invalid enumeration string:"+str);
				throw OsfMsAjaxFactory.msAjaxError.argument("str");
			}
			return parsed;
		},
		delayExecutionAndCache: function OSF_OUtil$delayExecutionAndCache() {
			var obj={ calc: arguments[0] };
			return function () {
				if (obj.calc) {
					obj.val=obj.calc.apply(this, arguments);
					delete obj.calc;
				}
				return obj.val;
			};
		},
		getUniqueId: function OSF_OUtil$getUniqueId() {
			_uniqueId=_uniqueId+1;
			return _uniqueId.toString();
		},
		formatString: function OSF_OUtil$formatString() {
			var args=arguments;
			var source=args[0];
			return source.replace(/{(\d+)}/gm, function (match, number) {
				var index=parseInt(number, 10)+1;
				return args[index]===undefined ? '{'+number+'}' : args[index];
			});
		},
		generateConversationId: function OSF_OUtil$generateConversationId() {
			return [_random(), _random(), (new Date()).getTime().toString()].join('_');
		},
		getFrameName: function OSF_OUtil$getFrameName(cacheKey) {
			return _xdmSessionKeyPrefix+cacheKey+this.generateConversationId();
		},
		addXdmInfoAsHash: function OSF_OUtil$addXdmInfoAsHash(url, xdmInfoValue) {
			return OSF.OUtil.addInfoAsHash(url, _xdmInfoKey, xdmInfoValue, false);
		},
		addSerializerVersionAsHash: function OSF_OUtil$addSerializerVersionAsHash(url, serializerVersion) {
			return OSF.OUtil.addInfoAsHash(url, _serializerVersionKey, serializerVersion, true);
		},
		addInfoAsHash: function OSF_OUtil$addInfoAsHash(url, keyName, infoValue, encodeInfo) {
			url=url.trim() || '';
			var urlParts=url.split(_fragmentSeparator);
			var urlWithoutFragment=urlParts.shift();
			var fragment=urlParts.join(_fragmentSeparator);
			var newFragment;
			if (encodeInfo) {
				newFragment=[keyName, encodeURIComponent(infoValue), fragment].join('');
			}
			else {
				newFragment=[fragment, keyName, infoValue].join('');
			}
			return [urlWithoutFragment, _fragmentSeparator, newFragment].join('');
		},
		parseHostInfoFromWindowName: function OSF_OUtil$parseHostInfoFromWindowName(skipSessionStorage, windowName) {
			return OSF.OUtil.parseInfoFromWindowName(skipSessionStorage, windowName, OSF.WindowNameItemKeys.HostInfo);
		},
		parseXdmInfo: function OSF_OUtil$parseXdmInfo(skipSessionStorage) {
			var xdmInfoValue=OSF.OUtil.parseXdmInfoWithGivenFragment(skipSessionStorage, window.location.hash);
			if (!xdmInfoValue) {
				xdmInfoValue=OSF.OUtil.parseXdmInfoFromWindowName(skipSessionStorage, window.name);
			}
			return xdmInfoValue;
		},
		parseXdmInfoFromWindowName: function OSF_OUtil$parseXdmInfoFromWindowName(skipSessionStorage, windowName) {
			return OSF.OUtil.parseInfoFromWindowName(skipSessionStorage, windowName, OSF.WindowNameItemKeys.XdmInfo);
		},
		parseXdmInfoWithGivenFragment: function OSF_OUtil$parseXdmInfoWithGivenFragment(skipSessionStorage, fragment) {
			return OSF.OUtil.parseInfoWithGivenFragment(_xdmInfoKey, _xdmSessionKeyPrefix, false, skipSessionStorage, fragment);
		},
		parseSerializerVersion: function OSF_OUtil$parseSerializerVersion(skipSessionStorage) {
			var serializerVersion=OSF.OUtil.parseSerializerVersionWithGivenFragment(skipSessionStorage, window.location.hash);
			if (isNaN(serializerVersion)) {
				serializerVersion=OSF.OUtil.parseSerializerVersionFromWindowName(skipSessionStorage, window.name);
			}
			return serializerVersion;
		},
		parseSerializerVersionFromWindowName: function OSF_OUtil$parseSerializerVersionFromWindowName(skipSessionStorage, windowName) {
			return parseInt(OSF.OUtil.parseInfoFromWindowName(skipSessionStorage, windowName, OSF.WindowNameItemKeys.SerializerVersion));
		},
		parseSerializerVersionWithGivenFragment: function OSF_OUtil$parseSerializerVersionWithGivenFragment(skipSessionStorage, fragment) {
			return parseInt(OSF.OUtil.parseInfoWithGivenFragment(_serializerVersionKey, _serializerVersionKeyPrefix, true, skipSessionStorage, fragment));
		},
		parseInfoFromWindowName: function OSF_OUtil$parseInfoFromWindowName(skipSessionStorage, windowName, infoKey) {
			try {
				var windowNameObj=JSON.parse(windowName);
				var infoValue=windowNameObj !=null ? windowNameObj[infoKey] : null;
				var osfSessionStorage=_getSessionStorage();
				if (!skipSessionStorage && osfSessionStorage && windowNameObj !=null) {
					var sessionKey=windowNameObj[OSF.WindowNameItemKeys.BaseFrameName]+infoKey;
					if (infoValue) {
						osfSessionStorage.setItem(sessionKey, infoValue);
					}
					else {
						infoValue=osfSessionStorage.getItem(sessionKey);
					}
				}
				return infoValue;
			}
			catch (Exception) {
				return null;
			}
		},
		parseInfoWithGivenFragment: function OSF_OUtil$parseInfoWithGivenFragment(infoKey, infoKeyPrefix, decodeInfo, skipSessionStorage, fragment) {
			var fragmentParts=fragment.split(infoKey);
			var infoValue=fragmentParts.length > 1 ? fragmentParts[fragmentParts.length - 1] : null;
			if (decodeInfo && infoValue !=null) {
				if (infoValue.indexOf(_fragmentInfoDelimiter) >=0) {
					infoValue=infoValue.split(_fragmentInfoDelimiter)[0];
				}
				infoValue=decodeURIComponent(infoValue);
			}
			var osfSessionStorage=_getSessionStorage();
			if (!skipSessionStorage && osfSessionStorage) {
				var sessionKeyStart=window.name.indexOf(infoKeyPrefix);
				if (sessionKeyStart > -1) {
					var sessionKeyEnd=window.name.indexOf(";", sessionKeyStart);
					if (sessionKeyEnd==-1) {
						sessionKeyEnd=window.name.length;
					}
					var sessionKey=window.name.substring(sessionKeyStart, sessionKeyEnd);
					if (infoValue) {
						osfSessionStorage.setItem(sessionKey, infoValue);
					}
					else {
						infoValue=osfSessionStorage.getItem(sessionKey);
					}
				}
			}
			return infoValue;
		},
		getConversationId: function OSF_OUtil$getConversationId() {
			var searchString=window.location.search;
			var conversationId=null;
			if (searchString) {
				var index=searchString.indexOf("&");
				conversationId=index > 0 ? searchString.substring(1, index) : searchString.substr(1);
				if (conversationId && conversationId.charAt(conversationId.length - 1)==='=') {
					conversationId=conversationId.substring(0, conversationId.length - 1);
					if (conversationId) {
						conversationId=decodeURIComponent(conversationId);
					}
				}
			}
			return conversationId;
		},
		getInfoItems: function OSF_OUtil$getInfoItems(strInfo) {
			var items=strInfo.split("$");
			if (typeof items[1]=="undefined") {
				items=strInfo.split("|");
			}
			if (typeof items[1]=="undefined") {
				items=strInfo.split("%7C");
			}
			return items;
		},
		getXdmFieldValue: function OSF_OUtil$getXdmFieldValue(xdmFieldName, skipSessionStorage) {
			var fieldValue='';
			var xdmInfoValue=OSF.OUtil.parseXdmInfo(skipSessionStorage);
			if (xdmInfoValue) {
				var items=OSF.OUtil.getInfoItems(xdmInfoValue);
				if (items !=undefined && items.length >=3) {
					switch (xdmFieldName) {
						case OSF.XdmFieldName.ConversationUrl:
							fieldValue=items[2];
							break;
						case OSF.XdmFieldName.AppId:
							fieldValue=items[1];
							break;
					}
				}
			}
			return fieldValue;
		},
		validateParamObject: function OSF_OUtil$validateParamObject(params, expectedProperties, callback) {
			var e=Function._validateParams(arguments, [{ name: "params", type: Object, mayBeNull: false },
				{ name: "expectedProperties", type: Object, mayBeNull: false },
				{ name: "callback", type: Function, mayBeNull: true }
			]);
			if (e)
				throw e;
			for (var p in expectedProperties) {
				e=Function._validateParameter(params[p], expectedProperties[p], p);
				if (e)
					throw e;
			}
		},
		writeProfilerMark: function OSF_OUtil$writeProfilerMark(text) {
			if (window.msWriteProfilerMark) {
				window.msWriteProfilerMark(text);
				OsfMsAjaxFactory.msAjaxDebug.trace(text);
			}
		},
		outputDebug: function OSF_OUtil$outputDebug(text) {
			if (typeof (OsfMsAjaxFactory) !=='undefined' && OsfMsAjaxFactory.msAjaxDebug && OsfMsAjaxFactory.msAjaxDebug.trace) {
				OsfMsAjaxFactory.msAjaxDebug.trace(text);
			}
		},
		defineNondefaultProperty: function OSF_OUtil$defineNondefaultProperty(obj, prop, descriptor, attributes) {
			descriptor=descriptor || {};
			for (var nd in attributes) {
				var attribute=attributes[nd];
				if (descriptor[attribute]==undefined) {
					descriptor[attribute]=true;
				}
			}
			Object.defineProperty(obj, prop, descriptor);
			return obj;
		},
		defineNondefaultProperties: function OSF_OUtil$defineNondefaultProperties(obj, descriptors, attributes) {
			descriptors=descriptors || {};
			for (var prop in descriptors) {
				OSF.OUtil.defineNondefaultProperty(obj, prop, descriptors[prop], attributes);
			}
			return obj;
		},
		defineEnumerableProperty: function OSF_OUtil$defineEnumerableProperty(obj, prop, descriptor) {
			return OSF.OUtil.defineNondefaultProperty(obj, prop, descriptor, ["enumerable"]);
		},
		defineEnumerableProperties: function OSF_OUtil$defineEnumerableProperties(obj, descriptors) {
			return OSF.OUtil.defineNondefaultProperties(obj, descriptors, ["enumerable"]);
		},
		defineMutableProperty: function OSF_OUtil$defineMutableProperty(obj, prop, descriptor) {
			return OSF.OUtil.defineNondefaultProperty(obj, prop, descriptor, ["writable", "enumerable", "configurable"]);
		},
		defineMutableProperties: function OSF_OUtil$defineMutableProperties(obj, descriptors) {
			return OSF.OUtil.defineNondefaultProperties(obj, descriptors, ["writable", "enumerable", "configurable"]);
		},
		finalizeProperties: function OSF_OUtil$finalizeProperties(obj, descriptor) {
			descriptor=descriptor || {};
			var props=Object.getOwnPropertyNames(obj);
			var propsLength=props.length;
			for (var i=0; i < propsLength; i++) {
				var prop=props[i];
				var desc=Object.getOwnPropertyDescriptor(obj, prop);
				if (!desc.get && !desc.set) {
					desc.writable=descriptor.writable || false;
				}
				desc.configurable=descriptor.configurable || false;
				desc.enumerable=descriptor.enumerable || true;
				Object.defineProperty(obj, prop, desc);
			}
			return obj;
		},
		mapList: function OSF_OUtil$MapList(list, mapFunction) {
			var ret=[];
			if (list) {
				for (var item in list) {
					ret.push(mapFunction(list[item]));
				}
			}
			return ret;
		},
		listContainsKey: function OSF_OUtil$listContainsKey(list, key) {
			for (var item in list) {
				if (key==item) {
					return true;
				}
			}
			return false;
		},
		listContainsValue: function OSF_OUtil$listContainsElement(list, value) {
			for (var item in list) {
				if (value==list[item]) {
					return true;
				}
			}
			return false;
		},
		augmentList: function OSF_OUtil$augmentList(list, addenda) {
			var add=list.push ? function (key, value) { list.push(value); } : function (key, value) { list[key]=value; };
			for (var key in addenda) {
				add(key, addenda[key]);
			}
		},
		redefineList: function OSF_Outil$redefineList(oldList, newList) {
			for (var key1 in oldList) {
				delete oldList[key1];
			}
			for (var key2 in newList) {
				oldList[key2]=newList[key2];
			}
		},
		isArray: function OSF_OUtil$isArray(obj) {
			return Object.prototype.toString.apply(obj)==="[object Array]";
		},
		isFunction: function OSF_OUtil$isFunction(obj) {
			return Object.prototype.toString.apply(obj)==="[object Function]";
		},
		isDate: function OSF_OUtil$isDate(obj) {
			return Object.prototype.toString.apply(obj)==="[object Date]";
		},
		addEventListener: function OSF_OUtil$addEventListener(element, eventName, listener) {
			if (element.addEventListener) {
				element.addEventListener(eventName, listener, false);
			}
			else if ((Sys.Browser.agent===Sys.Browser.InternetExplorer) && element.attachEvent) {
				element.attachEvent("on"+eventName, listener);
			}
			else {
				element["on"+eventName]=listener;
			}
		},
		removeEventListener: function OSF_OUtil$removeEventListener(element, eventName, listener) {
			if (element.removeEventListener) {
				element.removeEventListener(eventName, listener, false);
			}
			else if ((Sys.Browser.agent===Sys.Browser.InternetExplorer) && element.detachEvent) {
				element.detachEvent("on"+eventName, listener);
			}
			else {
				element["on"+eventName]=null;
			}
		},
		getCookieValue: function OSF_OUtil$getCookieValue(cookieName) {
			var tmpCookieString=RegExp(cookieName+"[^;]+").exec(document.cookie);
			return tmpCookieString.toString().replace(/^[^=]+./, "");
		},
		xhrGet: function OSF_OUtil$xhrGet(url, onSuccess, onError) {
			var xmlhttp;
			try {
				xmlhttp=new XMLHttpRequest();
				xmlhttp.onreadystatechange=function () {
					if (xmlhttp.readyState==4) {
						if (xmlhttp.status==200) {
							onSuccess(xmlhttp.responseText);
						}
						else {
							onError(xmlhttp.status);
						}
					}
				};
				xmlhttp.open("GET", url, true);
				xmlhttp.send();
			}
			catch (ex) {
				onError(ex);
			}
		},
		xhrGetFull: function OSF_OUtil$xhrGetFull(url, oneDriveFileName, onSuccess, onError) {
			var xmlhttp;
			var requestedFileName=oneDriveFileName;
			try {
				xmlhttp=new XMLHttpRequest();
				xmlhttp.onreadystatechange=function () {
					if (xmlhttp.readyState==4) {
						if (xmlhttp.status==200) {
							onSuccess(xmlhttp, requestedFileName);
						}
						else {
							onError(xmlhttp.status);
						}
					}
				};
				xmlhttp.open("GET", url, true);
				xmlhttp.send();
			}
			catch (ex) {
				onError(ex);
			}
		},
		encodeBase64: function OSF_Outil$encodeBase64(input) {
			if (!input)
				return input;
			var codex="ABCDEFGHIJKLMNOP"+"QRSTUVWXYZabcdef"+"ghijklmnopqrstuv"+"wxyz0123456789+/=";
			var output=[];
			var temp=[];
			var index=0;
			var c1, c2, c3, a, b, c;
			var i;
			var length=input.length;
			do {
				c1=input.charCodeAt(index++);
				c2=input.charCodeAt(index++);
				c3=input.charCodeAt(index++);
				i=0;
				a=c1 & 255;
				b=c1 >> 8;
				c=c2 & 255;
				temp[i++]=a >> 2;
				temp[i++]=((a & 3) << 4) | (b >> 4);
				temp[i++]=((b & 15) << 2) | (c >> 6);
				temp[i++]=c & 63;
				if (!isNaN(c2)) {
					a=c2 >> 8;
					b=c3 & 255;
					c=c3 >> 8;
					temp[i++]=a >> 2;
					temp[i++]=((a & 3) << 4) | (b >> 4);
					temp[i++]=((b & 15) << 2) | (c >> 6);
					temp[i++]=c & 63;
				}
				if (isNaN(c2)) {
					temp[i - 1]=64;
				}
				else if (isNaN(c3)) {
					temp[i - 2]=64;
					temp[i - 1]=64;
				}
				for (var t=0; t < i; t++) {
					output.push(codex.charAt(temp[t]));
				}
			} while (index < length);
			return output.join("");
		},
		getSessionStorage: function OSF_Outil$getSessionStorage() {
			return _getSessionStorage();
		},
		getLocalStorage: function OSF_Outil$getLocalStorage() {
			if (!_safeLocalStorage) {
				try {
					var localStorage=window.localStorage;
				}
				catch (ex) {
					localStorage=null;
				}
				_safeLocalStorage=new OfficeExt.SafeStorage(localStorage);
			}
			return _safeLocalStorage;
		},
		convertIntToCssHexColor: function OSF_Outil$convertIntToCssHexColor(val) {
			var hex="#"+(Number(val)+0x1000000).toString(16).slice(-6);
			return hex;
		},
		attachClickHandler: function OSF_Outil$attachClickHandler(element, handler) {
			element.onclick=function (e) {
				handler();
			};
			element.ontouchend=function (e) {
				handler();
				e.preventDefault();
			};
		},
		getQueryStringParamValue: function OSF_Outil$getQueryStringParamValue(queryString, paramName) {
			var e=Function._validateParams(arguments, [{ name: "queryString", type: String, mayBeNull: false },
				{ name: "paramName", type: String, mayBeNull: false }
			]);
			if (e) {
				OsfMsAjaxFactory.msAjaxDebug.trace("OSF_Outil_getQueryStringParamValue: Parameters cannot be null.");
				return "";
			}
			var queryExp=new RegExp("[\\?&]"+paramName+"=([^&#]*)", "i");
			if (!queryExp.test(queryString)) {
				OsfMsAjaxFactory.msAjaxDebug.trace("OSF_Outil_getQueryStringParamValue: The parameter is not found.");
				return "";
			}
			return queryExp.exec(queryString)[1];
		},
		isiOS: function OSF_Outil$isiOS() {
			return (window.navigator.userAgent.match(/(iPad|iPhone|iPod)/g) ? true : false);
		},
		isChrome: function OSF_Outil$isChrome() {
			return (window.navigator.userAgent.indexOf("Chrome") > 0) && !OSF.OUtil.isEdge();
		},
		isEdge: function OSF_Outil$isEdge() {
			return window.navigator.userAgent.indexOf("Edge") > 0;
		},
		isIE: function OSF_Outil$isIE() {
			return window.navigator.userAgent.indexOf("Trident") > 0;
		},
		isFirefox: function OSF_Outil$isFirefox() {
			return window.navigator.userAgent.indexOf("Firefox") > 0;
		},
		shallowCopy: function OSF_Outil$shallowCopy(sourceObj) {
			if (sourceObj==null) {
				return null;
			}
			else if (!(sourceObj instanceof Object)) {
				return sourceObj;
			}
			else if (Array.isArray(sourceObj)) {
				var copyArr=[];
				for (var i=0; i < sourceObj.length; i++) {
					copyArr.push(sourceObj[i]);
				}
				return copyArr;
			}
			else {
				var copyObj=sourceObj.constructor();
				for (var property in sourceObj) {
					if (sourceObj.hasOwnProperty(property)) {
						copyObj[property]=sourceObj[property];
					}
				}
				return copyObj;
			}
		},
		createObject: function OSF_Outil$createObject(properties) {
			var obj=null;
			if (properties) {
				obj={};
				var len=properties.length;
				for (var i=0; i < len; i++) {
					obj[properties[i].name]=properties[i].value;
				}
			}
			return obj;
		},
		addClass: function OSF_OUtil$addClass(elmt, val) {
			if (!OSF.OUtil.hasClass(elmt, val)) {
				var className=elmt.getAttribute(_classN);
				if (className) {
					elmt.setAttribute(_classN, className+" "+val);
				}
				else {
					elmt.setAttribute(_classN, val);
				}
			}
		},
		removeClass: function OSF_OUtil$removeClass(elmt, val) {
			if (OSF.OUtil.hasClass(elmt, val)) {
				var className=elmt.getAttribute(_classN);
				var reg=new RegExp('(\\s|^)'+val+'(\\s|$)');
				className=className.replace(reg, '');
				elmt.setAttribute(_classN, className);
			}
		},
		hasClass: function OSF_OUtil$hasClass(elmt, clsName) {
			var className=elmt.getAttribute(_classN);
			return className && className.match(new RegExp('(\\s|^)'+clsName+'(\\s|$)'));
		},
		focusToFirstTabbable: function OSF_OUtil$focusToFirstTabbable(all, backward) {
			var next;
			var focused=false;
			var candidate;
			var setFlag=function (e) {
				focused=true;
			};
			var findNextPos=function (allLen, currPos, backward) {
				if (currPos < 0 || currPos > allLen) {
					return -1;
				}
				else if (currPos===0 && backward) {
					return -1;
				}
				else if (currPos===allLen - 1 && !backward) {
					return -1;
				}
				if (backward) {
					return currPos - 1;
				}
				else {
					return currPos+1;
				}
			};
			all=_reOrderTabbableElements(all);
			next=backward ? all.length - 1 : 0;
			if (all.length===0) {
				return null;
			}
			while (!focused && next >=0 && next < all.length) {
				candidate=all[next];
				window.focus();
				candidate.addEventListener('focus', setFlag);
				candidate.focus();
				candidate.removeEventListener('focus', setFlag);
				next=findNextPos(all.length, next, backward);
				if (!focused && candidate===document.activeElement) {
					focused=true;
				}
			}
			if (focused) {
				return candidate;
			}
			else {
				return null;
			}
		},
		focusToNextTabbable: function OSF_OUtil$focusToNextTabbable(all, curr, shift) {
			var currPos;
			var next;
			var focused=false;
			var candidate;
			var setFlag=function (e) {
				focused=true;
			};
			var findCurrPos=function (all, curr) {
				var i=0;
				for (; i < all.length; i++) {
					if (all[i]===curr) {
						return i;
					}
				}
				return -1;
			};
			var findNextPos=function (allLen, currPos, shift) {
				if (currPos < 0 || currPos > allLen) {
					return -1;
				}
				else if (currPos===0 && shift) {
					return -1;
				}
				else if (currPos===allLen - 1 && !shift) {
					return -1;
				}
				if (shift) {
					return currPos - 1;
				}
				else {
					return currPos+1;
				}
			};
			all=_reOrderTabbableElements(all);
			currPos=findCurrPos(all, curr);
			next=findNextPos(all.length, currPos, shift);
			if (next < 0) {
				return null;
			}
			while (!focused && next >=0 && next < all.length) {
				candidate=all[next];
				candidate.addEventListener('focus', setFlag);
				candidate.focus();
				candidate.removeEventListener('focus', setFlag);
				next=findNextPos(all.length, next, shift);
				if (!focused && candidate===document.activeElement) {
					focused=true;
				}
			}
			if (focused) {
				return candidate;
			}
			else {
				return null;
			}
		}
	};
})();
OSF.OUtil.Guid=(function () {
	var hexCode=["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e", "f"];
	return {
		generateNewGuid: function OSF_Outil_Guid$generateNewGuid() {
			var result="";
			var tick=(new Date()).getTime();
			var index=0;
			for (; index < 32 && tick > 0; index++) {
				if (index==8 || index==12 || index==16 || index==20) {
					result+="-";
				}
				result+=hexCode[tick % 16];
				tick=Math.floor(tick / 16);
			}
			for (; index < 32; index++) {
				if (index==8 || index==12 || index==16 || index==20) {
					result+="-";
				}
				result+=hexCode[Math.floor(Math.random() * 16)];
			}
			return result;
		}
	};
})();
window.OSF=OSF;
OSF.OUtil.setNamespace("OSF", window);
OSF.MessageIDs={
	"FetchBundleUrl": 0,
	"LoadReactBundle": 1,
	"LoadBundleSuccess": 2,
	"LoadBundleError": 3
};
OSF.AppName={
	Unsupported: 0,
	Excel: 1,
	Word: 2,
	PowerPoint: 4,
	Outlook: 8,
	ExcelWebApp: 16,
	WordWebApp: 32,
	OutlookWebApp: 64,
	Project: 128,
	AccessWebApp: 256,
	PowerpointWebApp: 512,
	ExcelIOS: 1024,
	Sway: 2048,
	WordIOS: 4096,
	PowerPointIOS: 8192,
	Access: 16384,
	Lync: 32768,
	OutlookIOS: 65536,
	OneNoteWebApp: 131072,
	OneNote: 262144,
	ExcelWinRT: 524288,
	WordWinRT: 1048576,
	PowerpointWinRT: 2097152,
	OutlookAndroid: 4194304,
	OneNoteWinRT: 8388608,
	ExcelAndroid: 8388609,
	VisioWebApp: 8388610,
	OneNoteIOS: 8388611,
	WordAndroid: 8388613,
	PowerpointAndroid: 8388614
};
OSF.InternalPerfMarker={
	DataCoercionBegin: "Agave.HostCall.CoerceDataStart",
	DataCoercionEnd: "Agave.HostCall.CoerceDataEnd"
};
OSF.HostCallPerfMarker={
	IssueCall: "Agave.HostCall.IssueCall",
	ReceiveResponse: "Agave.HostCall.ReceiveResponse",
	RuntimeExceptionRaised: "Agave.HostCall.RuntimeExecptionRaised"
};
OSF.AgaveHostAction={
	"Select": 0,
	"UnSelect": 1,
	"CancelDialog": 2,
	"InsertAgave": 3,
	"CtrlF6In": 4,
	"CtrlF6Exit": 5,
	"CtrlF6ExitShift": 6,
	"SelectWithError": 7,
	"NotifyHostError": 8,
	"RefreshAddinCommands": 9,
	"PageIsReady": 10,
	"TabIn": 11,
	"TabInShift": 12,
	"TabExit": 13,
	"TabExitShift": 14,
	"EscExit": 15,
	"F2Exit": 16,
	"ExitNoFocusable": 17,
	"ExitNoFocusableShift": 18,
	"MouseEnter": 19,
	"MouseLeave": 20
};
OSF.SharedConstants={
	"NotificationConversationIdSuffix": '_ntf'
};
OSF.DialogMessageType={
	DialogMessageReceived: 0,
	DialogParentMessageReceived: 1,
	DialogClosed: 12006
};
OSF.OfficeAppContext=function OSF_OfficeAppContext(id, appName, appVersion, appUILocale, dataLocale, docUrl, clientMode, settings, reason, osfControlType, eToken, correlationId, appInstanceId, touchEnabled, commerceAllowed, appMinorVersion, requirementMatrix, hostCustomMessage, hostFullVersion, clientWindowHeight, clientWindowWidth, addinName, appDomains, dialogRequirementMatrix) {
	this._id=id;
	this._appName=appName;
	this._appVersion=appVersion;
	this._appUILocale=appUILocale;
	this._dataLocale=dataLocale;
	this._docUrl=docUrl;
	this._clientMode=clientMode;
	this._settings=settings;
	this._reason=reason;
	this._osfControlType=osfControlType;
	this._eToken=eToken;
	this._correlationId=correlationId;
	this._appInstanceId=appInstanceId;
	this._touchEnabled=touchEnabled;
	this._commerceAllowed=commerceAllowed;
	this._appMinorVersion=appMinorVersion;
	this._requirementMatrix=requirementMatrix;
	this._hostCustomMessage=hostCustomMessage;
	this._hostFullVersion=hostFullVersion;
	this._isDialog=false;
	this._clientWindowHeight=clientWindowHeight;
	this._clientWindowWidth=clientWindowWidth;
	this._addinName=addinName;
	this._appDomains=appDomains;
	this._dialogRequirementMatrix=dialogRequirementMatrix;
	this.get_id=function get_id() { return this._id; };
	this.get_appName=function get_appName() { return this._appName; };
	this.get_appVersion=function get_appVersion() { return this._appVersion; };
	this.get_appUILocale=function get_appUILocale() { return this._appUILocale; };
	this.get_dataLocale=function get_dataLocale() { return this._dataLocale; };
	this.get_docUrl=function get_docUrl() { return this._docUrl; };
	this.get_clientMode=function get_clientMode() { return this._clientMode; };
	this.get_bindings=function get_bindings() { return this._bindings; };
	this.get_settings=function get_settings() { return this._settings; };
	this.get_reason=function get_reason() { return this._reason; };
	this.get_osfControlType=function get_osfControlType() { return this._osfControlType; };
	this.get_eToken=function get_eToken() { return this._eToken; };
	this.get_correlationId=function get_correlationId() { return this._correlationId; };
	this.get_appInstanceId=function get_appInstanceId() { return this._appInstanceId; };
	this.get_touchEnabled=function get_touchEnabled() { return this._touchEnabled; };
	this.get_commerceAllowed=function get_commerceAllowed() { return this._commerceAllowed; };
	this.get_appMinorVersion=function get_appMinorVersion() { return this._appMinorVersion; };
	this.get_requirementMatrix=function get_requirementMatrix() { return this._requirementMatrix; };
	this.get_dialogRequirementMatrix=function get_dialogRequirementMatrix() { return this._dialogRequirementMatrix; };
	this.get_hostCustomMessage=function get_hostCustomMessage() { return this._hostCustomMessage; };
	this.get_hostFullVersion=function get_hostFullVersion() { return this._hostFullVersion; };
	this.get_isDialog=function get_isDialog() { return this._isDialog; };
	this.get_clientWindowHeight=function get_clientWindowHeight() { return this._clientWindowHeight; };
	this.get_clientWindowWidth=function get_clientWindowWidth() { return this._clientWindowWidth; };
	this.get_addinName=function get_addinName() { return this._addinName; };
	this.get_appDomains=function get_appDomains() { return this._appDomains; };
};
OSF.OsfControlType={
	DocumentLevel: 0,
	ContainerLevel: 1
};
OSF.ClientMode={
	ReadOnly: 0,
	ReadWrite: 1
};
OSF.OUtil.setNamespace("Microsoft", window);
OSF.OUtil.setNamespace("Office", Microsoft);
OSF.OUtil.setNamespace("Client", Microsoft.Office);
OSF.OUtil.setNamespace("WebExtension", Microsoft.Office);
Microsoft.Office.WebExtension.InitializationReason={
	Inserted: "inserted",
	DocumentOpened: "documentOpened"
};
Microsoft.Office.WebExtension.ValueFormat={
	Unformatted: "unformatted",
	Formatted: "formatted"
};
Microsoft.Office.WebExtension.FilterType={
	All: "all"
};
Microsoft.Office.WebExtension.PlatformType={
	PC: "PC",
	OfficeOnline: "OfficeOnline",
	Mac: "Mac",
	iOS: "iOS",
	Android: "Android",
	Universal: "Universal"
};
Microsoft.Office.WebExtension.HostType={
	Word: "Word",
	Excel: "Excel",
	PowerPoint: "PowerPoint",
	Outlook: "Outlook",
	OneNote: "OneNote",
	Project: "Project",
	Access: "Access"
};
Microsoft.Office.WebExtension.Parameters={
	BindingType: "bindingType",
	CoercionType: "coercionType",
	ValueFormat: "valueFormat",
	FilterType: "filterType",
	Columns: "columns",
	SampleData: "sampleData",
	GoToType: "goToType",
	SelectionMode: "selectionMode",
	Id: "id",
	PromptText: "promptText",
	ItemName: "itemName",
	FailOnCollision: "failOnCollision",
	StartRow: "startRow",
	StartColumn: "startColumn",
	RowCount: "rowCount",
	ColumnCount: "columnCount",
	Callback: "callback",
	AsyncContext: "asyncContext",
	Data: "data",
	Rows: "rows",
	OverwriteIfStale: "overwriteIfStale",
	FileType: "fileType",
	EventType: "eventType",
	Handler: "handler",
	SliceSize: "sliceSize",
	SliceIndex: "sliceIndex",
	ActiveView: "activeView",
	Status: "status",
	PlatformType: "platformType",
	HostType: "hostType",
	ForceConsent: "forceConsent",
	ForceAddAccount: "forceAddAccount",
	AuthChallenge: "authChallenge",
	Reserved: "reserved",
	Xml: "xml",
	Namespace: "namespace",
	Prefix: "prefix",
	XPath: "xPath",
	Text: "text",
	ImageLeft: "imageLeft",
	ImageTop: "imageTop",
	ImageWidth: "imageWidth",
	ImageHeight: "imageHeight",
	TaskId: "taskId",
	FieldId: "fieldId",
	FieldValue: "fieldValue",
	ServerUrl: "serverUrl",
	ListName: "listName",
	ResourceId: "resourceId",
	ViewType: "viewType",
	ViewName: "viewName",
	GetRawValue: "getRawValue",
	CellFormat: "cellFormat",
	TableOptions: "tableOptions",
	TaskIndex: "taskIndex",
	ResourceIndex: "resourceIndex",
	CustomFieldId: "customFieldId",
	Url: "url",
	MessageHandler: "messageHandler",
	Width: "width",
	Height: "height",
	RequireHTTPs: "requireHTTPS",
	MessageToParent: "messageToParent",
	DisplayInIframe: "displayInIframe",
	MessageContent: "messageContent",
	HideTitle: "hideTitle",
	UseDeviceIndependentPixels: "useDeviceIndependentPixels",
	AppCommandInvocationCompletedData: "appCommandInvocationCompletedData",
	Base64: "base64",
	FormId: "formId"
};
OSF.OUtil.setNamespace("DDA", OSF);
OSF.DDA.DocumentMode={
	ReadOnly: 1,
	ReadWrite: 0
};
OSF.DDA.PropertyDescriptors={
	AsyncResultStatus: "AsyncResultStatus"
};
OSF.DDA.EventDescriptors={};
OSF.DDA.ListDescriptors={};
OSF.DDA.UI={};
OSF.DDA.getXdmEventName=function OSF_DDA$GetXdmEventName(id, eventType) {
	if (eventType==Microsoft.Office.WebExtension.EventType.BindingSelectionChanged ||
		eventType==Microsoft.Office.WebExtension.EventType.BindingDataChanged ||
		eventType==Microsoft.Office.WebExtension.EventType.DataNodeDeleted ||
		eventType==Microsoft.Office.WebExtension.EventType.DataNodeInserted ||
		eventType==Microsoft.Office.WebExtension.EventType.DataNodeReplaced) {
		return id+"_"+eventType;
	}
	else {
		return eventType;
	}
};
OSF.DDA.MethodDispId={
	dispidMethodMin: 64,
	dispidGetSelectedDataMethod: 64,
	dispidSetSelectedDataMethod: 65,
	dispidAddBindingFromSelectionMethod: 66,
	dispidAddBindingFromPromptMethod: 67,
	dispidGetBindingMethod: 68,
	dispidReleaseBindingMethod: 69,
	dispidGetBindingDataMethod: 70,
	dispidSetBindingDataMethod: 71,
	dispidAddRowsMethod: 72,
	dispidClearAllRowsMethod: 73,
	dispidGetAllBindingsMethod: 74,
	dispidLoadSettingsMethod: 75,
	dispidSaveSettingsMethod: 76,
	dispidGetDocumentCopyMethod: 77,
	dispidAddBindingFromNamedItemMethod: 78,
	dispidAddColumnsMethod: 79,
	dispidGetDocumentCopyChunkMethod: 80,
	dispidReleaseDocumentCopyMethod: 81,
	dispidNavigateToMethod: 82,
	dispidGetActiveViewMethod: 83,
	dispidGetDocumentThemeMethod: 84,
	dispidGetOfficeThemeMethod: 85,
	dispidGetFilePropertiesMethod: 86,
	dispidClearFormatsMethod: 87,
	dispidSetTableOptionsMethod: 88,
	dispidSetFormatsMethod: 89,
	dispidExecuteRichApiRequestMethod: 93,
	dispidAppCommandInvocationCompletedMethod: 94,
	dispidCloseContainerMethod: 97,
	dispidGetAccessTokenMethod: 98,
	dispidOpenBrowserWindow: 102,
	dispidCreateDocumentMethod: 105,
	dispidInsertFormMethod: 106,
	dispidDisplayRibbonCalloutAsyncMethod: 109,
	dispidGetSelectedTaskMethod: 110,
	dispidGetSelectedResourceMethod: 111,
	dispidGetTaskMethod: 112,
	dispidGetResourceFieldMethod: 113,
	dispidGetWSSUrlMethod: 114,
	dispidGetTaskFieldMethod: 115,
	dispidGetProjectFieldMethod: 116,
	dispidGetSelectedViewMethod: 117,
	dispidGetTaskByIndexMethod: 118,
	dispidGetResourceByIndexMethod: 119,
	dispidSetTaskFieldMethod: 120,
	dispidSetResourceFieldMethod: 121,
	dispidGetMaxTaskIndexMethod: 122,
	dispidGetMaxResourceIndexMethod: 123,
	dispidCreateTaskMethod: 124,
	dispidAddDataPartMethod: 128,
	dispidGetDataPartByIdMethod: 129,
	dispidGetDataPartsByNamespaceMethod: 130,
	dispidGetDataPartXmlMethod: 131,
	dispidGetDataPartNodesMethod: 132,
	dispidDeleteDataPartMethod: 133,
	dispidGetDataNodeValueMethod: 134,
	dispidGetDataNodeXmlMethod: 135,
	dispidGetDataNodesMethod: 136,
	dispidSetDataNodeValueMethod: 137,
	dispidSetDataNodeXmlMethod: 138,
	dispidAddDataNamespaceMethod: 139,
	dispidGetDataUriByPrefixMethod: 140,
	dispidGetDataPrefixByUriMethod: 141,
	dispidGetDataNodeTextMethod: 142,
	dispidSetDataNodeTextMethod: 143,
	dispidMessageParentMethod: 144,
	dispidSendMessageMethod: 145,
	dispidMethodMax: 145
};
OSF.DDA.EventDispId={
	dispidEventMin: 0,
	dispidInitializeEvent: 0,
	dispidSettingsChangedEvent: 1,
	dispidDocumentSelectionChangedEvent: 2,
	dispidBindingSelectionChangedEvent: 3,
	dispidBindingDataChangedEvent: 4,
	dispidDocumentOpenEvent: 5,
	dispidDocumentCloseEvent: 6,
	dispidActiveViewChangedEvent: 7,
	dispidDocumentThemeChangedEvent: 8,
	dispidOfficeThemeChangedEvent: 9,
	dispidDialogMessageReceivedEvent: 10,
	dispidDialogNotificationShownInAddinEvent: 11,
	dispidDialogParentMessageReceivedEvent: 12,
	dispidObjectDeletedEvent: 13,
	dispidObjectSelectionChangedEvent: 14,
	dispidObjectDataChangedEvent: 15,
	dispidContentControlAddedEvent: 16,
	dispidActivationStatusChangedEvent: 32,
	dispidRichApiMessageEvent: 33,
	dispidAppCommandInvokedEvent: 39,
	dispidOlkItemSelectedChangedEvent: 46,
	dispidOlkRecipientsChangedEvent: 47,
	dispidOlkAppointmentTimeChangedEvent: 48,
	dispidOlkRecurrenceChangedEvent: 49,
	dispidTaskSelectionChangedEvent: 56,
	dispidResourceSelectionChangedEvent: 57,
	dispidViewSelectionChangedEvent: 58,
	dispidDataNodeAddedEvent: 60,
	dispidDataNodeReplacedEvent: 61,
	dispidDataNodeDeletedEvent: 62,
	dispidEventMax: 63
};
OSF.DDA.ErrorCodeManager=(function () {
	var _errorMappings={};
	return {
		getErrorArgs: function OSF_DDA_ErrorCodeManager$getErrorArgs(errorCode) {
			var errorArgs=_errorMappings[errorCode];
			if (!errorArgs) {
				errorArgs=_errorMappings[this.errorCodes.ooeInternalError];
			}
			else {
				if (!errorArgs.name) {
					errorArgs.name=_errorMappings[this.errorCodes.ooeInternalError].name;
				}
				if (!errorArgs.message) {
					errorArgs.message=_errorMappings[this.errorCodes.ooeInternalError].message;
				}
			}
			return errorArgs;
		},
		addErrorMessage: function OSF_DDA_ErrorCodeManager$addErrorMessage(errorCode, errorNameMessage) {
			_errorMappings[errorCode]=errorNameMessage;
		},
		errorCodes: {
			ooeSuccess: 0,
			ooeChunkResult: 1,
			ooeCoercionTypeNotSupported: 1000,
			ooeGetSelectionNotMatchDataType: 1001,
			ooeCoercionTypeNotMatchBinding: 1002,
			ooeInvalidGetRowColumnCounts: 1003,
			ooeSelectionNotSupportCoercionType: 1004,
			ooeInvalidGetStartRowColumn: 1005,
			ooeNonUniformPartialGetNotSupported: 1006,
			ooeGetDataIsTooLarge: 1008,
			ooeFileTypeNotSupported: 1009,
			ooeGetDataParametersConflict: 1010,
			ooeInvalidGetColumns: 1011,
			ooeInvalidGetRows: 1012,
			ooeInvalidReadForBlankRow: 1013,
			ooeUnsupportedDataObject: 2000,
			ooeCannotWriteToSelection: 2001,
			ooeDataNotMatchSelection: 2002,
			ooeOverwriteWorksheetData: 2003,
			ooeDataNotMatchBindingSize: 2004,
			ooeInvalidSetStartRowColumn: 2005,
			ooeInvalidDataFormat: 2006,
			ooeDataNotMatchCoercionType: 2007,
			ooeDataNotMatchBindingType: 2008,
			ooeSetDataIsTooLarge: 2009,
			ooeNonUniformPartialSetNotSupported: 2010,
			ooeInvalidSetColumns: 2011,
			ooeInvalidSetRows: 2012,
			ooeSetDataParametersConflict: 2013,
			ooeCellDataAmountBeyondLimits: 2014,
			ooeSelectionCannotBound: 3000,
			ooeBindingNotExist: 3002,
			ooeBindingToMultipleSelection: 3003,
			ooeInvalidSelectionForBindingType: 3004,
			ooeOperationNotSupportedOnThisBindingType: 3005,
			ooeNamedItemNotFound: 3006,
			ooeMultipleNamedItemFound: 3007,
			ooeInvalidNamedItemForBindingType: 3008,
			ooeUnknownBindingType: 3009,
			ooeOperationNotSupportedOnMatrixData: 3010,
			ooeInvalidColumnsForBinding: 3011,
			ooeSettingNameNotExist: 4000,
			ooeSettingsCannotSave: 4001,
			ooeSettingsAreStale: 4002,
			ooeOperationNotSupported: 5000,
			ooeInternalError: 5001,
			ooeDocumentReadOnly: 5002,
			ooeEventHandlerNotExist: 5003,
			ooeInvalidApiCallInContext: 5004,
			ooeShuttingDown: 5005,
			ooeUnsupportedEnumeration: 5007,
			ooeIndexOutOfRange: 5008,
			ooeBrowserAPINotSupported: 5009,
			ooeInvalidParam: 5010,
			ooeRequestTimeout: 5011,
			ooeInvalidOrTimedOutSession: 5012,
			ooeInvalidApiArguments: 5013,
			ooeOperationCancelled: 5014,
			ooeTooManyIncompleteRequests: 5100,
			ooeRequestTokenUnavailable: 5101,
			ooeActivityLimitReached: 5102,
			ooeCustomXmlNodeNotFound: 6000,
			ooeCustomXmlError: 6100,
			ooeCustomXmlExceedQuota: 6101,
			ooeCustomXmlOutOfDate: 6102,
			ooeNoCapability: 7000,
			ooeCannotNavTo: 7001,
			ooeSpecifiedIdNotExist: 7002,
			ooeNavOutOfBound: 7004,
			ooeElementMissing: 8000,
			ooeProtectedError: 8001,
			ooeInvalidCellsValue: 8010,
			ooeInvalidTableOptionValue: 8011,
			ooeInvalidFormatValue: 8012,
			ooeRowIndexOutOfRange: 8020,
			ooeColIndexOutOfRange: 8021,
			ooeFormatValueOutOfRange: 8022,
			ooeCellFormatAmountBeyondLimits: 8023,
			ooeMemoryFileLimit: 11000,
			ooeNetworkProblemRetrieveFile: 11001,
			ooeInvalidSliceSize: 11002,
			ooeInvalidCallback: 11101,
			ooeInvalidWidth: 12000,
			ooeInvalidHeight: 12001,
			ooeNavigationError: 12002,
			ooeInvalidScheme: 12003,
			ooeAppDomains: 12004,
			ooeRequireHTTPS: 12005,
			ooeWebDialogClosed: 12006,
			ooeDialogAlreadyOpened: 12007,
			ooeEndUserAllow: 12008,
			ooeEndUserIgnore: 12009,
			ooeNotUILessDialog: 12010,
			ooeCrossZone: 12011,
			ooeNotSSOAgave: 13000,
			ooeSSOUserNotSignedIn: 13001,
			ooeSSOUserAborted: 13002,
			ooeSSOUnsupportedUserIdentity: 13003,
			ooeSSOInvalidResourceUrl: 13004,
			ooeSSOInvalidGrant: 13005,
			ooeSSOClientError: 13006,
			ooeSSOServerError: 13007,
			ooeAddinIsAlreadyRequestingToken: 13008,
			ooeSSOUserConsentNotSupportedByCurrentAddinCategory: 13009,
			ooeSSOConnectionLost: 13010
		},
		initializeErrorMessages: function OSF_DDA_ErrorCodeManager$initializeErrorMessages(stringNS) {
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotSupported]={ name: stringNS.L_InvalidCoercion, message: stringNS.L_CoercionTypeNotSupported };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetSelectionNotMatchDataType]={ name: stringNS.L_DataReadError, message: stringNS.L_GetSelectionNotSupported };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotMatchBinding]={ name: stringNS.L_InvalidCoercion, message: stringNS.L_CoercionTypeNotMatchBinding };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetRowColumnCounts]={ name: stringNS.L_DataReadError, message: stringNS.L_InvalidGetRowColumnCounts };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSelectionNotSupportCoercionType]={ name: stringNS.L_DataReadError, message: stringNS.L_SelectionNotSupportCoercionType };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetStartRowColumn]={ name: stringNS.L_DataReadError, message: stringNS.L_InvalidGetStartRowColumn };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNonUniformPartialGetNotSupported]={ name: stringNS.L_DataReadError, message: stringNS.L_NonUniformPartialGetNotSupported };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetDataIsTooLarge]={ name: stringNS.L_DataReadError, message: stringNS.L_GetDataIsTooLarge };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeFileTypeNotSupported]={ name: stringNS.L_DataReadError, message: stringNS.L_FileTypeNotSupported };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetDataParametersConflict]={ name: stringNS.L_DataReadError, message: stringNS.L_GetDataParametersConflict };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetColumns]={ name: stringNS.L_DataReadError, message: stringNS.L_InvalidGetColumns };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetRows]={ name: stringNS.L_DataReadError, message: stringNS.L_InvalidGetRows };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidReadForBlankRow]={ name: stringNS.L_DataReadError, message: stringNS.L_InvalidReadForBlankRow };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedDataObject]={ name: stringNS.L_DataWriteError, message: stringNS.L_UnsupportedDataObject };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCannotWriteToSelection]={ name: stringNS.L_DataWriteError, message: stringNS.L_CannotWriteToSelection };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchSelection]={ name: stringNS.L_DataWriteError, message: stringNS.L_DataNotMatchSelection };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOverwriteWorksheetData]={ name: stringNS.L_DataWriteError, message: stringNS.L_OverwriteWorksheetData };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchBindingSize]={ name: stringNS.L_DataWriteError, message: stringNS.L_DataNotMatchBindingSize };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetStartRowColumn]={ name: stringNS.L_DataWriteError, message: stringNS.L_InvalidSetStartRowColumn };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidDataFormat]={ name: stringNS.L_InvalidFormat, message: stringNS.L_InvalidDataFormat };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchCoercionType]={ name: stringNS.L_InvalidDataObject, message: stringNS.L_DataNotMatchCoercionType };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchBindingType]={ name: stringNS.L_InvalidDataObject, message: stringNS.L_DataNotMatchBindingType };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSetDataIsTooLarge]={ name: stringNS.L_DataWriteError, message: stringNS.L_SetDataIsTooLarge };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNonUniformPartialSetNotSupported]={ name: stringNS.L_DataWriteError, message: stringNS.L_NonUniformPartialSetNotSupported };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetColumns]={ name: stringNS.L_DataWriteError, message: stringNS.L_InvalidSetColumns };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetRows]={ name: stringNS.L_DataWriteError, message: stringNS.L_InvalidSetRows };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSetDataParametersConflict]={ name: stringNS.L_DataWriteError, message: stringNS.L_SetDataParametersConflict };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSelectionCannotBound]={ name: stringNS.L_BindingCreationError, message: stringNS.L_SelectionCannotBound };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeBindingNotExist]={ name: stringNS.L_InvalidBindingError, message: stringNS.L_BindingNotExist };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeBindingToMultipleSelection]={ name: stringNS.L_BindingCreationError, message: stringNS.L_BindingToMultipleSelection };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSelectionForBindingType]={ name: stringNS.L_BindingCreationError, message: stringNS.L_InvalidSelectionForBindingType };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupportedOnThisBindingType]={ name: stringNS.L_InvalidBindingOperation, message: stringNS.L_OperationNotSupportedOnThisBindingType };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNamedItemNotFound]={ name: stringNS.L_BindingCreationError, message: stringNS.L_NamedItemNotFound };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeMultipleNamedItemFound]={ name: stringNS.L_BindingCreationError, message: stringNS.L_MultipleNamedItemFound };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidNamedItemForBindingType]={ name: stringNS.L_BindingCreationError, message: stringNS.L_InvalidNamedItemForBindingType };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnknownBindingType]={ name: stringNS.L_InvalidBinding, message: stringNS.L_UnknownBindingType };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupportedOnMatrixData]={ name: stringNS.L_InvalidBindingOperation, message: stringNS.L_OperationNotSupportedOnMatrixData };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidColumnsForBinding]={ name: stringNS.L_InvalidBinding, message: stringNS.L_InvalidColumnsForBinding };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingNameNotExist]={ name: stringNS.L_ReadSettingsError, message: stringNS.L_SettingNameNotExist };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingsCannotSave]={ name: stringNS.L_SaveSettingsError, message: stringNS.L_SettingsCannotSave };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingsAreStale]={ name: stringNS.L_SettingsStaleError, message: stringNS.L_SettingsAreStale };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupported]={ name: stringNS.L_HostError, message: stringNS.L_OperationNotSupported };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError]={ name: stringNS.L_InternalError, message: stringNS.L_InternalErrorDescription };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDocumentReadOnly]={ name: stringNS.L_PermissionDenied, message: stringNS.L_DocumentReadOnly };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerNotExist]={ name: stringNS.L_EventRegistrationError, message: stringNS.L_EventHandlerNotExist };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext]={ name: stringNS.L_InvalidAPICall, message: stringNS.L_InvalidApiCallInContext };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeShuttingDown]={ name: stringNS.L_ShuttingDown, message: stringNS.L_ShuttingDown };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedEnumeration]={ name: stringNS.L_UnsupportedEnumeration, message: stringNS.L_UnsupportedEnumerationMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeIndexOutOfRange]={ name: stringNS.L_IndexOutOfRange, message: stringNS.L_IndexOutOfRange };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeBrowserAPINotSupported]={ name: stringNS.L_APINotSupported, message: stringNS.L_BrowserAPINotSupported };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequestTimeout]={ name: stringNS.L_APICallFailed, message: stringNS.L_RequestTimeout };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidOrTimedOutSession]={ name: stringNS.L_InvalidOrTimedOutSession, message: stringNS.L_InvalidOrTimedOutSessionMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeTooManyIncompleteRequests]={ name: stringNS.L_APICallFailed, message: stringNS.L_TooManyIncompleteRequests };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequestTokenUnavailable]={ name: stringNS.L_APICallFailed, message: stringNS.L_RequestTokenUnavailable };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeActivityLimitReached]={ name: stringNS.L_APICallFailed, message: stringNS.L_ActivityLimitReached };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiArguments]={ name: stringNS.L_APICallFailed, message: stringNS.L_InvalidApiArgumentsMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlNodeNotFound]={ name: stringNS.L_InvalidNode, message: stringNS.L_CustomXmlNodeNotFound };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlError]={ name: stringNS.L_CustomXmlError, message: stringNS.L_CustomXmlError };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlExceedQuota]={ name: stringNS.L_CustomXmlExceedQuotaName, message: stringNS.L_CustomXmlExceedQuotaMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlOutOfDate]={ name: stringNS.L_CustomXmlOutOfDateName, message: stringNS.L_CustomXmlOutOfDateMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability]={ name: stringNS.L_PermissionDenied, message: stringNS.L_NoCapability };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCannotNavTo]={ name: stringNS.L_CannotNavigateTo, message: stringNS.L_CannotNavigateTo };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSpecifiedIdNotExist]={ name: stringNS.L_SpecifiedIdNotExist, message: stringNS.L_SpecifiedIdNotExist };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNavOutOfBound]={ name: stringNS.L_NavOutOfBound, message: stringNS.L_NavOutOfBound };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCellDataAmountBeyondLimits]={ name: stringNS.L_DataWriteReminder, message: stringNS.L_CellDataAmountBeyondLimits };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeElementMissing]={ name: stringNS.L_MissingParameter, message: stringNS.L_ElementMissing };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeProtectedError]={ name: stringNS.L_PermissionDenied, message: stringNS.L_NoCapability };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidCellsValue]={ name: stringNS.L_InvalidValue, message: stringNS.L_InvalidCellsValue };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidTableOptionValue]={ name: stringNS.L_InvalidValue, message: stringNS.L_InvalidTableOptionValue };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidFormatValue]={ name: stringNS.L_InvalidValue, message: stringNS.L_InvalidFormatValue };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeRowIndexOutOfRange]={ name: stringNS.L_OutOfRange, message: stringNS.L_RowIndexOutOfRange };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeColIndexOutOfRange]={ name: stringNS.L_OutOfRange, message: stringNS.L_ColIndexOutOfRange };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeFormatValueOutOfRange]={ name: stringNS.L_OutOfRange, message: stringNS.L_FormatValueOutOfRange };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCellFormatAmountBeyondLimits]={ name: stringNS.L_FormattingReminder, message: stringNS.L_CellFormatAmountBeyondLimits };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeMemoryFileLimit]={ name: stringNS.L_MemoryLimit, message: stringNS.L_CloseFileBeforeRetrieve };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNetworkProblemRetrieveFile]={ name: stringNS.L_NetworkProblem, message: stringNS.L_NetworkProblemRetrieveFile };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSliceSize]={ name: stringNS.L_InvalidValue, message: stringNS.L_SliceSizeNotSupported };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDialogAlreadyOpened]={ name: stringNS.L_DisplayDialogError, message: stringNS.L_DialogAlreadyOpened };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidWidth]={ name: stringNS.L_IndexOutOfRange, message: stringNS.L_IndexOutOfRange };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidHeight]={ name: stringNS.L_IndexOutOfRange, message: stringNS.L_IndexOutOfRange };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNavigationError]={ name: stringNS.L_DisplayDialogError, message: stringNS.L_NetworkProblem };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidScheme]={ name: stringNS.L_DialogNavigateError, message: stringNS.L_DialogInvalidScheme };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeAppDomains]={ name: stringNS.L_DisplayDialogError, message: stringNS.L_DialogAddressNotTrusted };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequireHTTPS]={ name: stringNS.L_DisplayDialogError, message: stringNS.L_DialogRequireHTTPS };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeEndUserIgnore]={ name: stringNS.L_DisplayDialogError, message: stringNS.L_UserClickIgnore };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCrossZone]={ name: stringNS.L_DisplayDialogError, message: stringNS.L_NewWindowCrossZoneErrorString };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNotSSOAgave]={ name: stringNS.L_APINotSupported, message: stringNS.L_InvalidSSOAddinMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUserNotSignedIn]={ name: stringNS.L_UserNotSignedIn, message: stringNS.L_UserNotSignedIn };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUserAborted]={ name: stringNS.L_UserAborted, message: stringNS.L_UserAbortedMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUnsupportedUserIdentity]={ name: stringNS.L_UnsupportedUserIdentity, message: stringNS.L_UnsupportedUserIdentityMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOInvalidResourceUrl]={ name: stringNS.L_InvalidResourceUrl, message: stringNS.L_InvalidResourceUrlMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOInvalidGrant]={ name: stringNS.L_InvalidGrant, message: stringNS.L_InvalidGrantMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOClientError]={ name: stringNS.L_SSOClientError, message: stringNS.L_SSOClientErrorMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOServerError]={ name: stringNS.L_SSOServerError, message: stringNS.L_SSOServerErrorMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeAddinIsAlreadyRequestingToken]={ name: stringNS.L_AddinIsAlreadyRequestingToken, message: stringNS.L_AddinIsAlreadyRequestingTokenMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUserConsentNotSupportedByCurrentAddinCategory]={ name: stringNS.L_SSOUserConsentNotSupportedByCurrentAddinCategory, message: stringNS.L_SSOUserConsentNotSupportedByCurrentAddinCategoryMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOConnectionLost]={ name: stringNS.L_SSOConnectionLostError, message: stringNS.L_SSOConnectionLostErrorMessage };
			_errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationCancelled]={ name: stringNS.L_OperationCancelledError, message: stringNS.L_OperationCancelledErrorMessage };
		}
	};
})();
var OfficeExt;
(function (OfficeExt) {
	var Requirement;
	(function (Requirement) {
		var RequirementVersion=(function () {
			function RequirementVersion() {
			}
			return RequirementVersion;
		})();
		Requirement.RequirementVersion=RequirementVersion;
		var RequirementMatrix=(function () {
			function RequirementMatrix(_setMap) {
				this.isSetSupported=function _isSetSupported(name, minVersion) {
					if (name==undefined) {
						return false;
					}
					if (minVersion==undefined) {
						minVersion=0;
					}
					var setSupportArray=this._setMap;
					var sets=setSupportArray._sets;
					if (sets.hasOwnProperty(name.toLowerCase())) {
						var setMaxVersion=sets[name.toLowerCase()];
						try {
							var setMaxVersionNum=this._getVersion(setMaxVersion);
							minVersion=minVersion+"";
							var minVersionNum=this._getVersion(minVersion);
							if (setMaxVersionNum.major > 0 && setMaxVersionNum.major > minVersionNum.major) {
								return true;
							}
							if (setMaxVersionNum.minor > 0 &&
								setMaxVersionNum.minor > 0 &&
								setMaxVersionNum.major==minVersionNum.major &&
								setMaxVersionNum.minor >=minVersionNum.minor) {
								return true;
							}
						}
						catch (e) {
							return false;
						}
					}
					return false;
				};
				this._getVersion=function (version) {
					var temp=version.split(".");
					var major=0;
					var minor=0;
					if (temp.length < 2 && isNaN(Number(version))) {
						throw "version format incorrect";
					}
					else {
						major=Number(temp[0]);
						if (temp.length >=2) {
							minor=Number(temp[1]);
						}
						if (isNaN(major) || isNaN(minor)) {
							throw "version format incorrect";
						}
					}
					var result={ "minor": minor, "major": major };
					return result;
				};
				this._setMap=_setMap;
				this.isSetSupported=this.isSetSupported.bind(this);
			}
			return RequirementMatrix;
		})();
		Requirement.RequirementMatrix=RequirementMatrix;
		var DefaultSetRequirement=(function () {
			function DefaultSetRequirement(setMap) {
				this._addSetMap=function DefaultSetRequirement_addSetMap(addedSet) {
					for (var name in addedSet) {
						this._sets[name]=addedSet[name];
					}
				};
				this._sets=setMap;
			}
			return DefaultSetRequirement;
		})();
		Requirement.DefaultSetRequirement=DefaultSetRequirement;
		var DefaultDialogSetRequirement=(function (_super) {
			__extends(DefaultDialogSetRequirement, _super);
			function DefaultDialogSetRequirement() {
				_super.call(this, {
					"dialogapi": 1.1
				});
			}
			return DefaultDialogSetRequirement;
		})(DefaultSetRequirement);
		Requirement.DefaultDialogSetRequirement=DefaultDialogSetRequirement;
		var ExcelClientDefaultSetRequirement=(function (_super) {
			__extends(ExcelClientDefaultSetRequirement, _super);
			function ExcelClientDefaultSetRequirement() {
				_super.call(this, {
					"bindingevents": 1.1,
					"documentevents": 1.1,
					"excelapi": 1.1,
					"matrixbindings": 1.1,
					"matrixcoercion": 1.1,
					"selection": 1.1,
					"settings": 1.1,
					"tablebindings": 1.1,
					"tablecoercion": 1.1,
					"textbindings": 1.1,
					"textcoercion": 1.1
				});
			}
			return ExcelClientDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.ExcelClientDefaultSetRequirement=ExcelClientDefaultSetRequirement;
		var ExcelClientV1DefaultSetRequirement=(function (_super) {
			__extends(ExcelClientV1DefaultSetRequirement, _super);
			function ExcelClientV1DefaultSetRequirement() {
				_super.call(this);
				this._addSetMap({
					"imagecoercion": 1.1
				});
			}
			return ExcelClientV1DefaultSetRequirement;
		})(ExcelClientDefaultSetRequirement);
		Requirement.ExcelClientV1DefaultSetRequirement=ExcelClientV1DefaultSetRequirement;
		var OutlookClientDefaultSetRequirement=(function (_super) {
			__extends(OutlookClientDefaultSetRequirement, _super);
			function OutlookClientDefaultSetRequirement() {
				_super.call(this, {
					"mailbox": 1.3
				});
			}
			return OutlookClientDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.OutlookClientDefaultSetRequirement=OutlookClientDefaultSetRequirement;
		var WordClientDefaultSetRequirement=(function (_super) {
			__extends(WordClientDefaultSetRequirement, _super);
			function WordClientDefaultSetRequirement() {
				_super.call(this, {
					"bindingevents": 1.1,
					"compressedfile": 1.1,
					"customxmlparts": 1.1,
					"documentevents": 1.1,
					"file": 1.1,
					"htmlcoercion": 1.1,
					"matrixbindings": 1.1,
					"matrixcoercion": 1.1,
					"ooxmlcoercion": 1.1,
					"pdffile": 1.1,
					"selection": 1.1,
					"settings": 1.1,
					"tablebindings": 1.1,
					"tablecoercion": 1.1,
					"textbindings": 1.1,
					"textcoercion": 1.1,
					"textfile": 1.1,
					"wordapi": 1.1
				});
			}
			return WordClientDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.WordClientDefaultSetRequirement=WordClientDefaultSetRequirement;
		var WordClientV1DefaultSetRequirement=(function (_super) {
			__extends(WordClientV1DefaultSetRequirement, _super);
			function WordClientV1DefaultSetRequirement() {
				_super.call(this);
				this._addSetMap({
					"customxmlparts": 1.2,
					"wordapi": 1.2,
					"imagecoercion": 1.1
				});
			}
			return WordClientV1DefaultSetRequirement;
		})(WordClientDefaultSetRequirement);
		Requirement.WordClientV1DefaultSetRequirement=WordClientV1DefaultSetRequirement;
		var PowerpointClientDefaultSetRequirement=(function (_super) {
			__extends(PowerpointClientDefaultSetRequirement, _super);
			function PowerpointClientDefaultSetRequirement() {
				_super.call(this, {
					"activeview": 1.1,
					"compressedfile": 1.1,
					"documentevents": 1.1,
					"file": 1.1,
					"pdffile": 1.1,
					"selection": 1.1,
					"settings": 1.1,
					"textcoercion": 1.1
				});
			}
			return PowerpointClientDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.PowerpointClientDefaultSetRequirement=PowerpointClientDefaultSetRequirement;
		var PowerpointClientV1DefaultSetRequirement=(function (_super) {
			__extends(PowerpointClientV1DefaultSetRequirement, _super);
			function PowerpointClientV1DefaultSetRequirement() {
				_super.call(this);
				this._addSetMap({
					"imagecoercion": 1.1
				});
			}
			return PowerpointClientV1DefaultSetRequirement;
		})(PowerpointClientDefaultSetRequirement);
		Requirement.PowerpointClientV1DefaultSetRequirement=PowerpointClientV1DefaultSetRequirement;
		var ProjectClientDefaultSetRequirement=(function (_super) {
			__extends(ProjectClientDefaultSetRequirement, _super);
			function ProjectClientDefaultSetRequirement() {
				_super.call(this, {
					"selection": 1.1,
					"textcoercion": 1.1
				});
			}
			return ProjectClientDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.ProjectClientDefaultSetRequirement=ProjectClientDefaultSetRequirement;
		var ExcelWebDefaultSetRequirement=(function (_super) {
			__extends(ExcelWebDefaultSetRequirement, _super);
			function ExcelWebDefaultSetRequirement() {
				_super.call(this, {
					"bindingevents": 1.1,
					"documentevents": 1.1,
					"matrixbindings": 1.1,
					"matrixcoercion": 1.1,
					"selection": 1.1,
					"settings": 1.1,
					"tablebindings": 1.1,
					"tablecoercion": 1.1,
					"textbindings": 1.1,
					"textcoercion": 1.1,
					"file": 1.1
				});
			}
			return ExcelWebDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.ExcelWebDefaultSetRequirement=ExcelWebDefaultSetRequirement;
		var WordWebDefaultSetRequirement=(function (_super) {
			__extends(WordWebDefaultSetRequirement, _super);
			function WordWebDefaultSetRequirement() {
				_super.call(this, {
					"compressedfile": 1.1,
					"documentevents": 1.1,
					"file": 1.1,
					"imagecoercion": 1.1,
					"matrixcoercion": 1.1,
					"ooxmlcoercion": 1.1,
					"pdffile": 1.1,
					"selection": 1.1,
					"settings": 1.1,
					"tablecoercion": 1.1,
					"textcoercion": 1.1,
					"textfile": 1.1
				});
			}
			return WordWebDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.WordWebDefaultSetRequirement=WordWebDefaultSetRequirement;
		var PowerpointWebDefaultSetRequirement=(function (_super) {
			__extends(PowerpointWebDefaultSetRequirement, _super);
			function PowerpointWebDefaultSetRequirement() {
				_super.call(this, {
					"activeview": 1.1,
					"settings": 1.1
				});
			}
			return PowerpointWebDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.PowerpointWebDefaultSetRequirement=PowerpointWebDefaultSetRequirement;
		var OutlookWebDefaultSetRequirement=(function (_super) {
			__extends(OutlookWebDefaultSetRequirement, _super);
			function OutlookWebDefaultSetRequirement() {
				_super.call(this, {
					"mailbox": 1.3
				});
			}
			return OutlookWebDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.OutlookWebDefaultSetRequirement=OutlookWebDefaultSetRequirement;
		var SwayWebDefaultSetRequirement=(function (_super) {
			__extends(SwayWebDefaultSetRequirement, _super);
			function SwayWebDefaultSetRequirement() {
				_super.call(this, {
					"activeview": 1.1,
					"documentevents": 1.1,
					"selection": 1.1,
					"settings": 1.1,
					"textcoercion": 1.1
				});
			}
			return SwayWebDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.SwayWebDefaultSetRequirement=SwayWebDefaultSetRequirement;
		var AccessWebDefaultSetRequirement=(function (_super) {
			__extends(AccessWebDefaultSetRequirement, _super);
			function AccessWebDefaultSetRequirement() {
				_super.call(this, {
					"bindingevents": 1.1,
					"partialtablebindings": 1.1,
					"settings": 1.1,
					"tablebindings": 1.1,
					"tablecoercion": 1.1
				});
			}
			return AccessWebDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.AccessWebDefaultSetRequirement=AccessWebDefaultSetRequirement;
		var ExcelIOSDefaultSetRequirement=(function (_super) {
			__extends(ExcelIOSDefaultSetRequirement, _super);
			function ExcelIOSDefaultSetRequirement() {
				_super.call(this, {
					"bindingevents": 1.1,
					"documentevents": 1.1,
					"matrixbindings": 1.1,
					"matrixcoercion": 1.1,
					"selection": 1.1,
					"settings": 1.1,
					"tablebindings": 1.1,
					"tablecoercion": 1.1,
					"textbindings": 1.1,
					"textcoercion": 1.1
				});
			}
			return ExcelIOSDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.ExcelIOSDefaultSetRequirement=ExcelIOSDefaultSetRequirement;
		var WordIOSDefaultSetRequirement=(function (_super) {
			__extends(WordIOSDefaultSetRequirement, _super);
			function WordIOSDefaultSetRequirement() {
				_super.call(this, {
					"bindingevents": 1.1,
					"compressedfile": 1.1,
					"customxmlparts": 1.1,
					"documentevents": 1.1,
					"file": 1.1,
					"htmlcoercion": 1.1,
					"matrixbindings": 1.1,
					"matrixcoercion": 1.1,
					"ooxmlcoercion": 1.1,
					"pdffile": 1.1,
					"selection": 1.1,
					"settings": 1.1,
					"tablebindings": 1.1,
					"tablecoercion": 1.1,
					"textbindings": 1.1,
					"textcoercion": 1.1,
					"textfile": 1.1
				});
			}
			return WordIOSDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.WordIOSDefaultSetRequirement=WordIOSDefaultSetRequirement;
		var WordIOSV1DefaultSetRequirement=(function (_super) {
			__extends(WordIOSV1DefaultSetRequirement, _super);
			function WordIOSV1DefaultSetRequirement() {
				_super.call(this);
				this._addSetMap({
					"customxmlparts": 1.2,
					"wordapi": 1.2
				});
			}
			return WordIOSV1DefaultSetRequirement;
		})(WordIOSDefaultSetRequirement);
		Requirement.WordIOSV1DefaultSetRequirement=WordIOSV1DefaultSetRequirement;
		var PowerpointIOSDefaultSetRequirement=(function (_super) {
			__extends(PowerpointIOSDefaultSetRequirement, _super);
			function PowerpointIOSDefaultSetRequirement() {
				_super.call(this, {
					"activeview": 1.1,
					"compressedfile": 1.1,
					"documentevents": 1.1,
					"file": 1.1,
					"pdffile": 1.1,
					"selection": 1.1,
					"settings": 1.1,
					"textcoercion": 1.1
				});
			}
			return PowerpointIOSDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.PowerpointIOSDefaultSetRequirement=PowerpointIOSDefaultSetRequirement;
		var OutlookIOSDefaultSetRequirement=(function (_super) {
			__extends(OutlookIOSDefaultSetRequirement, _super);
			function OutlookIOSDefaultSetRequirement() {
				_super.call(this, {
					"mailbox": 1.1
				});
			}
			return OutlookIOSDefaultSetRequirement;
		})(DefaultSetRequirement);
		Requirement.OutlookIOSDefaultSetRequirement=OutlookIOSDefaultSetRequirement;
		var RequirementsMatrixFactory=(function () {
			function RequirementsMatrixFactory() {
			}
			RequirementsMatrixFactory.initializeOsfDda=function () {
				OSF.OUtil.setNamespace("Requirement", OSF.DDA);
			};
			RequirementsMatrixFactory.getDefaultRequirementMatrix=function (appContext) {
				this.initializeDefaultSetMatrix();
				var defaultRequirementMatrix=undefined;
				var clientRequirement=appContext.get_requirementMatrix();
				if (clientRequirement !=undefined && clientRequirement.length > 0 && typeof (JSON) !=="undefined") {
					var matrixItem=JSON.parse(appContext.get_requirementMatrix().toLowerCase());
					defaultRequirementMatrix=new RequirementMatrix(new DefaultSetRequirement(matrixItem));
				}
				else {
					var appLocator=RequirementsMatrixFactory.getClientFullVersionString(appContext);
					if (RequirementsMatrixFactory.DefaultSetArrayMatrix !=undefined && RequirementsMatrixFactory.DefaultSetArrayMatrix[appLocator] !=undefined) {
						defaultRequirementMatrix=new RequirementMatrix(RequirementsMatrixFactory.DefaultSetArrayMatrix[appLocator]);
					}
					else {
						defaultRequirementMatrix=new RequirementMatrix(new DefaultSetRequirement({}));
					}
				}
				return defaultRequirementMatrix;
			};
			RequirementsMatrixFactory.getDefaultDialogRequirementMatrix=function (appContext) {
				var defaultRequirementMatrix=undefined;
				var clientRequirement=appContext.get_dialogRequirementMatrix();
				if (clientRequirement !=undefined && clientRequirement.length > 0 && typeof (JSON) !=="undefined") {
					var matrixItem=JSON.parse(appContext.get_requirementMatrix().toLowerCase());
					defaultRequirementMatrix=new RequirementMatrix(new DefaultSetRequirement(matrixItem));
				}
				else {
					defaultRequirementMatrix=new RequirementMatrix(new DefaultDialogSetRequirement());
				}
				return defaultRequirementMatrix;
			};
			RequirementsMatrixFactory.getClientFullVersionString=function (appContext) {
				var appMinorVersion=appContext.get_appMinorVersion();
				var appMinorVersionString="";
				var appFullVersion="";
				var appName=appContext.get_appName();
				var isIOSClient=appName==1024 ||
					appName==4096 ||
					appName==8192 ||
					appName==65536;
				if (isIOSClient && appContext.get_appVersion()==1) {
					if (appName==4096 && appMinorVersion >=15) {
						appFullVersion="16.00.01";
					}
					else {
						appFullVersion="16.00";
					}
				}
				else if (appContext.get_appName()==64) {
					appFullVersion=appContext.get_appVersion();
				}
				else {
					if (appMinorVersion < 10) {
						appMinorVersionString="0"+appMinorVersion;
					}
					else {
						appMinorVersionString=""+appMinorVersion;
					}
					appFullVersion=appContext.get_appVersion()+"."+appMinorVersionString;
				}
				return appContext.get_appName()+"-"+appFullVersion;
			};
			RequirementsMatrixFactory.initializeDefaultSetMatrix=function () {
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Excel_RCLIENT_1600]=new ExcelClientDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Word_RCLIENT_1600]=new WordClientDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.PowerPoint_RCLIENT_1600]=new PowerpointClientDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Excel_RCLIENT_1601]=new ExcelClientV1DefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Word_RCLIENT_1601]=new WordClientV1DefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.PowerPoint_RCLIENT_1601]=new PowerpointClientV1DefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Outlook_RCLIENT_1600]=new OutlookClientDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Excel_WAC_1600]=new ExcelWebDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Word_WAC_1600]=new WordWebDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Outlook_WAC_1600]=new OutlookWebDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Outlook_WAC_1601]=new OutlookWebDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Project_RCLIENT_1600]=new ProjectClientDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Access_WAC_1600]=new AccessWebDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.PowerPoint_WAC_1600]=new PowerpointWebDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Excel_IOS_1600]=new ExcelIOSDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.SWAY_WAC_1600]=new SwayWebDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Word_IOS_1600]=new WordIOSDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Word_IOS_16001]=new WordIOSV1DefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.PowerPoint_IOS_1600]=new PowerpointIOSDefaultSetRequirement();
				RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Outlook_IOS_1600]=new OutlookIOSDefaultSetRequirement();
			};
			RequirementsMatrixFactory.Excel_RCLIENT_1600="1-16.00";
			RequirementsMatrixFactory.Excel_RCLIENT_1601="1-16.01";
			RequirementsMatrixFactory.Word_RCLIENT_1600="2-16.00";
			RequirementsMatrixFactory.Word_RCLIENT_1601="2-16.01";
			RequirementsMatrixFactory.PowerPoint_RCLIENT_1600="4-16.00";
			RequirementsMatrixFactory.PowerPoint_RCLIENT_1601="4-16.01";
			RequirementsMatrixFactory.Outlook_RCLIENT_1600="8-16.00";
			RequirementsMatrixFactory.Excel_WAC_1600="16-16.00";
			RequirementsMatrixFactory.Word_WAC_1600="32-16.00";
			RequirementsMatrixFactory.Outlook_WAC_1600="64-16.00";
			RequirementsMatrixFactory.Outlook_WAC_1601="64-16.01";
			RequirementsMatrixFactory.Project_RCLIENT_1600="128-16.00";
			RequirementsMatrixFactory.Access_WAC_1600="256-16.00";
			RequirementsMatrixFactory.PowerPoint_WAC_1600="512-16.00";
			RequirementsMatrixFactory.Excel_IOS_1600="1024-16.00";
			RequirementsMatrixFactory.SWAY_WAC_1600="2048-16.00";
			RequirementsMatrixFactory.Word_IOS_1600="4096-16.00";
			RequirementsMatrixFactory.Word_IOS_16001="4096-16.00.01";
			RequirementsMatrixFactory.PowerPoint_IOS_1600="8192-16.00";
			RequirementsMatrixFactory.Outlook_IOS_1600="65536-16.00";
			RequirementsMatrixFactory.DefaultSetArrayMatrix={};
			return RequirementsMatrixFactory;
		})();
		Requirement.RequirementsMatrixFactory=RequirementsMatrixFactory;
	})(Requirement=OfficeExt.Requirement || (OfficeExt.Requirement={}));
})(OfficeExt || (OfficeExt={}));
OfficeExt.Requirement.RequirementsMatrixFactory.initializeOsfDda();
var OfficeExt;
(function (OfficeExt) {
	var HostName;
	(function (HostName) {
		var Host=(function () {
			function Host() {
				this.getDiagnostics=function _getDiagnostics(version) {
					var diagnostics={
						host: this.getHost(),
						version: (version || this.getDefaultVersion()),
						platform: this.getPlatform()
					};
					return diagnostics;
				};
				this.platformRemappings={
					web: Microsoft.Office.WebExtension.PlatformType.OfficeOnline,
					winrt: Microsoft.Office.WebExtension.PlatformType.Universal,
					win32: Microsoft.Office.WebExtension.PlatformType.PC,
					mac: Microsoft.Office.WebExtension.PlatformType.Mac,
					ios: Microsoft.Office.WebExtension.PlatformType.iOS,
					android: Microsoft.Office.WebExtension.PlatformType.Android
				};
				this.camelCaseMappings={
					powerpoint: Microsoft.Office.WebExtension.HostType.PowerPoint,
					onenote: Microsoft.Office.WebExtension.HostType.OneNote
				};
				this.hostInfo=OSF._OfficeAppFactory.getHostInfo();
				this.getHost=this.getHost.bind(this);
				this.getPlatform=this.getPlatform.bind(this);
				this.getDiagnostics=this.getDiagnostics.bind(this);
			}
			Host.prototype.capitalizeFirstLetter=function (input) {
				if (input) {
					return (input[0].toUpperCase()+input.slice(1).toLowerCase());
				}
				return input;
			};
			Host.getInstance=function () {
				if (Host.hostObj===undefined) {
					Host.hostObj=new Host();
				}
				return Host.hostObj;
			};
			Host.prototype.getPlatform=function () {
				if (this.hostInfo.hostPlatform) {
					var hostPlatform=this.hostInfo.hostPlatform.toLowerCase();
					if (this.platformRemappings[hostPlatform]) {
						return this.platformRemappings[hostPlatform];
					}
				}
				return null;
			};
			Host.prototype.getHost=function () {
				if (this.hostInfo.hostType) {
					var hostType=this.hostInfo.hostType.toLowerCase();
					if (this.camelCaseMappings[hostType]) {
						return this.camelCaseMappings[hostType];
					}
					hostType=this.capitalizeFirstLetter(this.hostInfo.hostType);
					if (Microsoft.Office.WebExtension.HostType[hostType]) {
						return Microsoft.Office.WebExtension.HostType[hostType];
					}
				}
				return null;
			};
			Host.prototype.getDefaultVersion=function () {
				if (this.getHost()) {
					return "16.0.0000.0000";
				}
				return null;
			};
			return Host;
		})();
		HostName.Host=Host;
	})(HostName=OfficeExt.HostName || (OfficeExt.HostName={}));
})(OfficeExt || (OfficeExt={}));
Microsoft.Office.WebExtension.ApplicationMode={
	WebEditor: "webEditor",
	WebViewer: "webViewer",
	Client: "client"
};
Microsoft.Office.WebExtension.DocumentMode={
	ReadOnly: "readOnly",
	ReadWrite: "readWrite"
};
OSF.NamespaceManager=(function OSF_NamespaceManager() {
	var _userOffice;
	var _useShortcut=false;
	return {
		enableShortcut: function OSF_NamespaceManager$enableShortcut() {
			if (!_useShortcut) {
				if (window.Office) {
					_userOffice=window.Office;
				}
				else {
					OSF.OUtil.setNamespace("Office", window);
				}
				window.Office=Microsoft.Office.WebExtension;
				_useShortcut=true;
			}
		},
		disableShortcut: function OSF_NamespaceManager$disableShortcut() {
			if (_useShortcut) {
				if (_userOffice) {
					window.Office=_userOffice;
				}
				else {
					OSF.OUtil.unsetNamespace("Office", window);
				}
				_useShortcut=false;
			}
		}
	};
})();
OSF.NamespaceManager.enableShortcut();
Microsoft.Office.WebExtension.useShortNamespace=function Microsoft_Office_WebExtension_useShortNamespace(useShortcut) {
	if (useShortcut) {
		OSF.NamespaceManager.enableShortcut();
	}
	else {
		OSF.NamespaceManager.disableShortcut();
	}
};
Microsoft.Office.WebExtension.select=function Microsoft_Office_WebExtension_select(str, errorCallback) {
	var promise;
	if (str && typeof str=="string") {
		var index=str.indexOf("#");
		if (index !=-1) {
			var op=str.substring(0, index);
			var target=str.substring(index+1);
			switch (op) {
				case "binding":
				case "bindings":
					if (target) {
						promise=new OSF.DDA.BindingPromise(target);
					}
					break;
			}
		}
	}
	if (!promise) {
		if (errorCallback) {
			var callbackType=typeof errorCallback;
			if (callbackType=="function") {
				var callArgs={};
				callArgs[Microsoft.Office.WebExtension.Parameters.Callback]=errorCallback;
				OSF.DDA.issueAsyncResult(callArgs, OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext, OSF.DDA.ErrorCodeManager.getErrorArgs(OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext));
			}
			else {
				throw OSF.OUtil.formatString(Strings.OfficeOM.L_CallbackNotAFunction, callbackType);
			}
		}
	}
	else {
		promise.onFail=errorCallback;
		return promise;
	}
};
OSF.DDA.Context=function OSF_DDA_Context(officeAppContext, document, license, appOM, getOfficeTheme) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"contentLanguage": {
			value: officeAppContext.get_dataLocale()
		},
		"displayLanguage": {
			value: officeAppContext.get_appUILocale()
		},
		"touchEnabled": {
			value: officeAppContext.get_touchEnabled()
		},
		"commerceAllowed": {
			value: officeAppContext.get_commerceAllowed()
		},
		"host": {
			value: OfficeExt.HostName.Host.getInstance().getHost()
		},
		"platform": {
			value: OfficeExt.HostName.Host.getInstance().getPlatform()
		},
		"diagnostics": {
			value: OfficeExt.HostName.Host.getInstance().getDiagnostics(officeAppContext.get_hostFullVersion())
		}
	});
	if (license) {
		OSF.OUtil.defineEnumerableProperty(this, "license", {
			value: license
		});
	}
	if (officeAppContext.ui) {
		OSF.OUtil.defineEnumerableProperty(this, "ui", {
			value: officeAppContext.ui
		});
	}
	if (officeAppContext.auth) {
		OSF.OUtil.defineEnumerableProperty(this, "auth", {
			value: officeAppContext.auth
		});
	}
	if (officeAppContext.application) {
		OSF.OUtil.defineEnumerableProperty(this, "application", {
			value: officeAppContext.application
		});
	}
	if (officeAppContext.get_isDialog()) {
		var requirements=OfficeExt.Requirement.RequirementsMatrixFactory.getDefaultDialogRequirementMatrix(officeAppContext);
		OSF.OUtil.defineEnumerableProperty(this, "requirements", {
			value: requirements
		});
	}
	else {
		if (document) {
			OSF.OUtil.defineEnumerableProperty(this, "document", {
				value: document
			});
		}
		if (appOM) {
			var displayName=appOM.displayName || "appOM";
			delete appOM.displayName;
			OSF.OUtil.defineEnumerableProperty(this, displayName, {
				value: appOM
			});
		}
		if (getOfficeTheme) {
			OSF.OUtil.defineEnumerableProperty(this, "officeTheme", {
				get: function () {
					return getOfficeTheme();
				}
			});
		}
		var requirements=OfficeExt.Requirement.RequirementsMatrixFactory.getDefaultRequirementMatrix(officeAppContext);
		OSF.OUtil.defineEnumerableProperty(this, "requirements", {
			value: requirements
		});
	}
};
OSF.DDA.OutlookContext=function OSF_DDA_OutlookContext(appContext, settings, license, appOM, getOfficeTheme) {
	OSF.DDA.OutlookContext.uber.constructor.call(this, appContext, null, license, appOM, getOfficeTheme);
	if (settings) {
		OSF.OUtil.defineEnumerableProperty(this, "roamingSettings", {
			value: settings
		});
	}
};
OSF.OUtil.extend(OSF.DDA.OutlookContext, OSF.DDA.Context);
OSF.DDA.OutlookAppOm=function OSF_DDA_OutlookAppOm(appContext, window, appReady) { };
OSF.DDA.Application=function OSF_DDA_Application(officeAppContext) {
};
OSF.DDA.Document=function OSF_DDA_Document(officeAppContext, settings) {
	var mode;
	switch (officeAppContext.get_clientMode()) {
		case OSF.ClientMode.ReadOnly:
			mode=Microsoft.Office.WebExtension.DocumentMode.ReadOnly;
			break;
		case OSF.ClientMode.ReadWrite:
			mode=Microsoft.Office.WebExtension.DocumentMode.ReadWrite;
			break;
	}
	;
	if (settings) {
		OSF.OUtil.defineEnumerableProperty(this, "settings", {
			value: settings
		});
	}
	;
	OSF.OUtil.defineMutableProperties(this, {
		"mode": {
			value: mode
		},
		"url": {
			value: officeAppContext.get_docUrl()
		}
	});
};
OSF.DDA.JsomDocument=function OSF_DDA_JsomDocument(officeAppContext, bindingFacade, settings) {
	OSF.DDA.JsomDocument.uber.constructor.call(this, officeAppContext, settings);
	if (bindingFacade) {
		OSF.OUtil.defineEnumerableProperty(this, "bindings", {
			get: function OSF_DDA_Document$GetBindings() { return bindingFacade; }
		});
	}
	var am=OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(this, [
		am.GetSelectedDataAsync,
		am.SetSelectedDataAsync
	]);
	OSF.DDA.DispIdHost.addEventSupport(this, new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged]));
};
OSF.OUtil.extend(OSF.DDA.JsomDocument, OSF.DDA.Document);
OSF.OUtil.defineEnumerableProperty(Microsoft.Office.WebExtension, "context", {
	get: function Microsoft_Office_WebExtension$GetContext() {
		var context;
		if (OSF && OSF._OfficeAppFactory) {
			context=OSF._OfficeAppFactory.getContext();
		}
		return context;
	}
});
OSF.DDA.License=function OSF_DDA_License(eToken) {
	OSF.OUtil.defineEnumerableProperty(this, "value", {
		value: eToken
	});
};
OSF.DDA.ApiMethodCall=function OSF_DDA_ApiMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, checkCallArgs, displayName) {
	var requiredCount=requiredParameters.length;
	var getInvalidParameterString=OSF.OUtil.delayExecutionAndCache(function () {
		return OSF.OUtil.formatString(Strings.OfficeOM.L_InvalidParameters, displayName);
	});
	this.verifyArguments=function OSF_DDA_ApiMethodCall$VerifyArguments(params, args) {
		for (var name in params) {
			var param=params[name];
			var arg=args[name];
			if (param["enum"]) {
				switch (typeof arg) {
					case "string":
						if (OSF.OUtil.listContainsValue(param["enum"], arg)) {
							break;
						}
					case "undefined":
						throw OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedEnumeration;
					default:
						throw getInvalidParameterString();
				}
			}
			if (param["types"]) {
				if (!OSF.OUtil.listContainsValue(param["types"], typeof arg)) {
					throw getInvalidParameterString();
				}
			}
		}
	};
	this.extractRequiredArguments=function OSF_DDA_ApiMethodCall$ExtractRequiredArguments(userArgs, caller, stateInfo) {
		if (userArgs.length < requiredCount) {
			throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_MissingRequiredArguments);
		}
		var requiredArgs=[];
		var index;
		for (index=0; index < requiredCount; index++) {
			requiredArgs.push(userArgs[index]);
		}
		this.verifyArguments(requiredParameters, requiredArgs);
		var ret={};
		for (index=0; index < requiredCount; index++) {
			var param=requiredParameters[index];
			var arg=requiredArgs[index];
			if (param.verify) {
				var isValid=param.verify(arg, caller, stateInfo);
				if (!isValid) {
					throw getInvalidParameterString();
				}
			}
			ret[param.name]=arg;
		}
		return ret;
	},
		this.fillOptions=function OSF_DDA_ApiMethodCall$FillOptions(options, requiredArgs, caller, stateInfo) {
			options=options || {};
			for (var optionName in supportedOptions) {
				if (!OSF.OUtil.listContainsKey(options, optionName)) {
					var value=undefined;
					var option=supportedOptions[optionName];
					if (option.calculate && requiredArgs) {
						value=option.calculate(requiredArgs, caller, stateInfo);
					}
					if (!value && option.defaultValue !==undefined) {
						value=option.defaultValue;
					}
					options[optionName]=value;
				}
			}
			return options;
		};
	this.constructCallArgs=function OSF_DAA_ApiMethodCall$ConstructCallArgs(required, options, caller, stateInfo) {
		var callArgs={};
		for (var r in required) {
			callArgs[r]=required[r];
		}
		for (var o in options) {
			callArgs[o]=options[o];
		}
		for (var s in privateStateCallbacks) {
			callArgs[s]=privateStateCallbacks[s](caller, stateInfo);
		}
		if (checkCallArgs) {
			callArgs=checkCallArgs(callArgs, caller, stateInfo);
		}
		return callArgs;
	};
};
OSF.OUtil.setNamespace("AsyncResultEnum", OSF.DDA);
OSF.DDA.AsyncResultEnum.Properties={
	Context: "Context",
	Value: "Value",
	Status: "Status",
	Error: "Error"
};
Microsoft.Office.WebExtension.AsyncResultStatus={
	Succeeded: "succeeded",
	Failed: "failed"
};
OSF.DDA.AsyncResultEnum.ErrorCode={
	Success: 0,
	Failed: 1
};
OSF.DDA.AsyncResultEnum.ErrorProperties={
	Name: "Name",
	Message: "Message",
	Code: "Code"
};
OSF.DDA.AsyncMethodNames={};
OSF.DDA.AsyncMethodNames.addNames=function (methodNames) {
	for (var entry in methodNames) {
		var am={};
		OSF.OUtil.defineEnumerableProperties(am, {
			"id": {
				value: entry
			},
			"displayName": {
				value: methodNames[entry]
			}
		});
		OSF.DDA.AsyncMethodNames[entry]=am;
	}
};
OSF.DDA.AsyncMethodCall=function OSF_DDA_AsyncMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, onSucceeded, onFailed, checkCallArgs, displayName) {
	var requiredCount=requiredParameters.length;
	var apiMethods=new OSF.DDA.ApiMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, checkCallArgs, displayName);
	function OSF_DAA_AsyncMethodCall$ExtractOptions(userArgs, requiredArgs, caller, stateInfo) {
		if (userArgs.length > requiredCount+2) {
			throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyArguments);
		}
		var options, parameterCallback;
		for (var i=userArgs.length - 1; i >=requiredCount; i--) {
			var argument=userArgs[i];
			switch (typeof argument) {
				case "object":
					if (options) {
						throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyOptionalObjects);
					}
					else {
						options=argument;
					}
					break;
				case "function":
					if (parameterCallback) {
						throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyOptionalFunction);
					}
					else {
						parameterCallback=argument;
					}
					break;
				default:
					throw OsfMsAjaxFactory.msAjaxError.argument(Strings.OfficeOM.L_InValidOptionalArgument);
					break;
			}
		}
		options=apiMethods.fillOptions(options, requiredArgs, caller, stateInfo);
		if (parameterCallback) {
			if (options[Microsoft.Office.WebExtension.Parameters.Callback]) {
				throw Strings.OfficeOM.L_RedundantCallbackSpecification;
			}
			else {
				options[Microsoft.Office.WebExtension.Parameters.Callback]=parameterCallback;
			}
		}
		apiMethods.verifyArguments(supportedOptions, options);
		return options;
	}
	;
	this.verifyAndExtractCall=function OSF_DAA_AsyncMethodCall$VerifyAndExtractCall(userArgs, caller, stateInfo) {
		var required=apiMethods.extractRequiredArguments(userArgs, caller, stateInfo);
		var options=OSF_DAA_AsyncMethodCall$ExtractOptions(userArgs, required, caller, stateInfo);
		var callArgs=apiMethods.constructCallArgs(required, options, caller, stateInfo);
		return callArgs;
	};
	this.processResponse=function OSF_DAA_AsyncMethodCall$ProcessResponse(status, response, caller, callArgs) {
		var payload;
		if (status==OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
			if (onSucceeded) {
				payload=onSucceeded(response, caller, callArgs);
			}
			else {
				payload=response;
			}
		}
		else {
			if (onFailed) {
				payload=onFailed(status, response);
			}
			else {
				payload=OSF.DDA.ErrorCodeManager.getErrorArgs(status);
			}
		}
		return payload;
	};
	this.getCallArgs=function (suppliedArgs) {
		var options, parameterCallback;
		for (var i=suppliedArgs.length - 1; i >=requiredCount; i--) {
			var argument=suppliedArgs[i];
			switch (typeof argument) {
				case "object":
					options=argument;
					break;
				case "function":
					parameterCallback=argument;
					break;
			}
		}
		options=options || {};
		if (parameterCallback) {
			options[Microsoft.Office.WebExtension.Parameters.Callback]=parameterCallback;
		}
		return options;
	};
};
OSF.DDA.AsyncMethodCallFactory=(function () {
	return {
		manufacture: function (params) {
			var supportedOptions=params.supportedOptions ? OSF.OUtil.createObject(params.supportedOptions) : [];
			var privateStateCallbacks=params.privateStateCallbacks ? OSF.OUtil.createObject(params.privateStateCallbacks) : [];
			return new OSF.DDA.AsyncMethodCall(params.requiredArguments || [], supportedOptions, privateStateCallbacks, params.onSucceeded, params.onFailed, params.checkCallArgs, params.method.displayName);
		}
	};
})();
OSF.DDA.AsyncMethodCalls={};
OSF.DDA.AsyncMethodCalls.define=function (callDefinition) {
	OSF.DDA.AsyncMethodCalls[callDefinition.method.id]=OSF.DDA.AsyncMethodCallFactory.manufacture(callDefinition);
};
OSF.DDA.Error=function OSF_DDA_Error(name, message, code) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"name": {
			value: name
		},
		"message": {
			value: message
		},
		"code": {
			value: code
		}
	});
};
OSF.DDA.AsyncResult=function OSF_DDA_AsyncResult(initArgs, errorArgs) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"value": {
			value: initArgs[OSF.DDA.AsyncResultEnum.Properties.Value]
		},
		"status": {
			value: errorArgs ? Microsoft.Office.WebExtension.AsyncResultStatus.Failed : Microsoft.Office.WebExtension.AsyncResultStatus.Succeeded
		}
	});
	if (initArgs[OSF.DDA.AsyncResultEnum.Properties.Context]) {
		OSF.OUtil.defineEnumerableProperty(this, "asyncContext", {
			value: initArgs[OSF.DDA.AsyncResultEnum.Properties.Context]
		});
	}
	if (errorArgs) {
		OSF.OUtil.defineEnumerableProperty(this, "error", {
			value: new OSF.DDA.Error(errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name], errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message], errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Code])
		});
	}
};
OSF.DDA.issueAsyncResult=function OSF_DDA$IssueAsyncResult(callArgs, status, payload) {
	var callback=callArgs[Microsoft.Office.WebExtension.Parameters.Callback];
	if (callback) {
		var asyncInitArgs={};
		asyncInitArgs[OSF.DDA.AsyncResultEnum.Properties.Context]=callArgs[Microsoft.Office.WebExtension.Parameters.AsyncContext];
		var errorArgs;
		if (status==OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
			asyncInitArgs[OSF.DDA.AsyncResultEnum.Properties.Value]=payload;
		}
		else {
			errorArgs={};
			payload=payload || OSF.DDA.ErrorCodeManager.getErrorArgs(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);
			errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Code]=status || OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
			errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name]=payload.name || payload;
			errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message]=payload.message || payload;
		}
		callback(new OSF.DDA.AsyncResult(asyncInitArgs, errorArgs));
	}
};
OSF.DDA.SyncMethodNames={};
OSF.DDA.SyncMethodNames.addNames=function (methodNames) {
	for (var entry in methodNames) {
		var am={};
		OSF.OUtil.defineEnumerableProperties(am, {
			"id": {
				value: entry
			},
			"displayName": {
				value: methodNames[entry]
			}
		});
		OSF.DDA.SyncMethodNames[entry]=am;
	}
};
OSF.DDA.SyncMethodCall=function OSF_DDA_SyncMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, checkCallArgs, displayName) {
	var requiredCount=requiredParameters.length;
	var apiMethods=new OSF.DDA.ApiMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, checkCallArgs, displayName);
	function OSF_DAA_SyncMethodCall$ExtractOptions(userArgs, requiredArgs, caller, stateInfo) {
		if (userArgs.length > requiredCount+1) {
			throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyArguments);
		}
		var options, parameterCallback;
		for (var i=userArgs.length - 1; i >=requiredCount; i--) {
			var argument=userArgs[i];
			switch (typeof argument) {
				case "object":
					if (options) {
						throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyOptionalObjects);
					}
					else {
						options=argument;
					}
					break;
				default:
					throw OsfMsAjaxFactory.msAjaxError.argument(Strings.OfficeOM.L_InValidOptionalArgument);
					break;
			}
		}
		options=apiMethods.fillOptions(options, requiredArgs, caller, stateInfo);
		apiMethods.verifyArguments(supportedOptions, options);
		return options;
	}
	;
	this.verifyAndExtractCall=function OSF_DAA_AsyncMethodCall$VerifyAndExtractCall(userArgs, caller, stateInfo) {
		var required=apiMethods.extractRequiredArguments(userArgs, caller, stateInfo);
		var options=OSF_DAA_SyncMethodCall$ExtractOptions(userArgs, required, caller, stateInfo);
		var callArgs=apiMethods.constructCallArgs(required, options, caller, stateInfo);
		return callArgs;
	};
};
OSF.DDA.SyncMethodCallFactory=(function () {
	return {
		manufacture: function (params) {
			var supportedOptions=params.supportedOptions ? OSF.OUtil.createObject(params.supportedOptions) : [];
			return new OSF.DDA.SyncMethodCall(params.requiredArguments || [], supportedOptions, params.privateStateCallbacks, params.checkCallArgs, params.method.displayName);
		}
	};
})();
OSF.DDA.SyncMethodCalls={};
OSF.DDA.SyncMethodCalls.define=function (callDefinition) {
	OSF.DDA.SyncMethodCalls[callDefinition.method.id]=OSF.DDA.SyncMethodCallFactory.manufacture(callDefinition);
};
OSF.DDA.ListType=(function () {
	var listTypes={};
	return {
		setListType: function OSF_DDA_ListType$AddListType(t, prop) { listTypes[t]=prop; },
		isListType: function OSF_DDA_ListType$IsListType(t) { return OSF.OUtil.listContainsKey(listTypes, t); },
		getDescriptor: function OSF_DDA_ListType$getDescriptor(t) { return listTypes[t]; }
	};
})();
OSF.DDA.HostParameterMap=function (specialProcessor, mappings) {
	var toHostMap="toHost";
	var fromHostMap="fromHost";
	var sourceData="sourceData";
	var self="self";
	var dynamicTypes={};
	dynamicTypes[Microsoft.Office.WebExtension.Parameters.Data]={
		toHost: function (data) {
			if (data !=null && data.rows !==undefined) {
				var tableData={};
				tableData[OSF.DDA.TableDataProperties.TableRows]=data.rows;
				tableData[OSF.DDA.TableDataProperties.TableHeaders]=data.headers;
				data=tableData;
			}
			return data;
		},
		fromHost: function (args) {
			return args;
		}
	};
	dynamicTypes[Microsoft.Office.WebExtension.Parameters.SampleData]=dynamicTypes[Microsoft.Office.WebExtension.Parameters.Data];
	function mapValues(preimageSet, mapping) {
		var ret=preimageSet ? {} : undefined;
		for (var entry in preimageSet) {
			var preimage=preimageSet[entry];
			var image;
			if (OSF.DDA.ListType.isListType(entry)) {
				image=[];
				for (var subEntry in preimage) {
					image.push(mapValues(preimage[subEntry], mapping));
				}
			}
			else if (OSF.OUtil.listContainsKey(dynamicTypes, entry)) {
				image=dynamicTypes[entry][mapping](preimage);
			}
			else if (mapping==fromHostMap && specialProcessor.preserveNesting(entry)) {
				image=mapValues(preimage, mapping);
			}
			else {
				var maps=mappings[entry];
				if (maps) {
					var map=maps[mapping];
					if (map) {
						image=map[preimage];
						if (image===undefined) {
							image=preimage;
						}
					}
				}
				else {
					image=preimage;
				}
			}
			ret[entry]=image;
		}
		return ret;
	}
	;
	function generateArguments(imageSet, parameters) {
		var ret;
		for (var param in parameters) {
			var arg;
			if (specialProcessor.isComplexType(param)) {
				arg=generateArguments(imageSet, mappings[param][toHostMap]);
			}
			else {
				arg=imageSet[param];
			}
			if (arg !=undefined) {
				if (!ret) {
					ret={};
				}
				var index=parameters[param];
				if (index==self) {
					index=param;
				}
				ret[index]=specialProcessor.pack(param, arg);
			}
		}
		return ret;
	}
	;
	function extractArguments(source, parameters, extracted) {
		if (!extracted) {
			extracted={};
		}
		for (var param in parameters) {
			var index=parameters[param];
			var value;
			if (index==self) {
				value=source;
			}
			else if (index==sourceData) {
				extracted[param]=source.toArray();
				continue;
			}
			else {
				value=source[index];
			}
			if (value===null || value===undefined) {
				extracted[param]=undefined;
			}
			else {
				value=specialProcessor.unpack(param, value);
				var map;
				if (specialProcessor.isComplexType(param)) {
					map=mappings[param][fromHostMap];
					if (specialProcessor.preserveNesting(param)) {
						extracted[param]=extractArguments(value, map);
					}
					else {
						extractArguments(value, map, extracted);
					}
				}
				else {
					if (OSF.DDA.ListType.isListType(param)) {
						map={};
						var entryDescriptor=OSF.DDA.ListType.getDescriptor(param);
						map[entryDescriptor]=self;
						var extractedValues=new Array(value.length);
						for (var item in value) {
							extractedValues[item]=extractArguments(value[item], map);
						}
						extracted[param]=extractedValues;
					}
					else {
						extracted[param]=value;
					}
				}
			}
		}
		return extracted;
	}
	;
	function applyMap(mapName, preimage, mapping) {
		var parameters=mappings[mapName][mapping];
		var image;
		if (mapping=="toHost") {
			var imageSet=mapValues(preimage, mapping);
			image=generateArguments(imageSet, parameters);
		}
		else if (mapping=="fromHost") {
			var argumentSet=extractArguments(preimage, parameters);
			image=mapValues(argumentSet, mapping);
		}
		return image;
	}
	;
	if (!mappings) {
		mappings={};
	}
	this.addMapping=function (mapName, description) {
		var toHost, fromHost;
		if (description.map) {
			toHost=description.map;
			fromHost={};
			for (var preimage in toHost) {
				var image=toHost[preimage];
				if (image==self) {
					image=preimage;
				}
				fromHost[image]=preimage;
			}
		}
		else {
			toHost=description.toHost;
			fromHost=description.fromHost;
		}
		var pair=mappings[mapName];
		if (pair) {
			var currMap=pair[toHostMap];
			for (var th in currMap)
				toHost[th]=currMap[th];
			currMap=pair[fromHostMap];
			for (var fh in currMap)
				fromHost[fh]=currMap[fh];
		}
		else {
			pair=mappings[mapName]={};
		}
		pair[toHostMap]=toHost;
		pair[fromHostMap]=fromHost;
	};
	this.toHost=function (mapName, preimage) { return applyMap(mapName, preimage, toHostMap); };
	this.fromHost=function (mapName, image) { return applyMap(mapName, image, fromHostMap); };
	this.self=self;
	this.sourceData=sourceData;
	this.addComplexType=function (ct) { specialProcessor.addComplexType(ct); };
	this.getDynamicType=function (dt) { return specialProcessor.getDynamicType(dt); };
	this.setDynamicType=function (dt, handler) { specialProcessor.setDynamicType(dt, handler); };
	this.dynamicTypes=dynamicTypes;
	this.doMapValues=function (preimageSet, mapping) { return mapValues(preimageSet, mapping); };
};
OSF.DDA.SpecialProcessor=function (complexTypes, dynamicTypes) {
	this.addComplexType=function OSF_DDA_SpecialProcessor$addComplexType(ct) {
		complexTypes.push(ct);
	};
	this.getDynamicType=function OSF_DDA_SpecialProcessor$getDynamicType(dt) {
		return dynamicTypes[dt];
	};
	this.setDynamicType=function OSF_DDA_SpecialProcessor$setDynamicType(dt, handler) {
		dynamicTypes[dt]=handler;
	};
	this.isComplexType=function OSF_DDA_SpecialProcessor$isComplexType(t) {
		return OSF.OUtil.listContainsValue(complexTypes, t);
	};
	this.isDynamicType=function OSF_DDA_SpecialProcessor$isDynamicType(p) {
		return OSF.OUtil.listContainsKey(dynamicTypes, p);
	};
	this.preserveNesting=function OSF_DDA_SpecialProcessor$preserveNesting(p) {
		var pn=[];
		if (OSF.DDA.PropertyDescriptors)
			pn.push(OSF.DDA.PropertyDescriptors.Subset);
		if (OSF.DDA.DataNodeEventProperties) {
			pn=pn.concat([
				OSF.DDA.DataNodeEventProperties.OldNode,
				OSF.DDA.DataNodeEventProperties.NewNode,
				OSF.DDA.DataNodeEventProperties.NextSiblingNode
			]);
		}
		return OSF.OUtil.listContainsValue(pn, p);
	};
	this.pack=function OSF_DDA_SpecialProcessor$pack(param, arg) {
		var value;
		if (this.isDynamicType(param)) {
			value=dynamicTypes[param].toHost(arg);
		}
		else {
			value=arg;
		}
		return value;
	};
	this.unpack=function OSF_DDA_SpecialProcessor$unpack(param, arg) {
		var value;
		if (this.isDynamicType(param)) {
			value=dynamicTypes[param].fromHost(arg);
		}
		else {
			value=arg;
		}
		return value;
	};
};
OSF.DDA.getDecoratedParameterMap=function (specialProcessor, initialDefs) {
	var parameterMap=new OSF.DDA.HostParameterMap(specialProcessor);
	var self=parameterMap.self;
	function createObject(properties) {
		var obj=null;
		if (properties) {
			obj={};
			var len=properties.length;
			for (var i=0; i < len; i++) {
				obj[properties[i].name]=properties[i].value;
			}
		}
		return obj;
	}
	parameterMap.define=function define(definition) {
		var args={};
		var toHost=createObject(definition.toHost);
		if (definition.invertible) {
			args.map=toHost;
		}
		else if (definition.canonical) {
			args.toHost=args.fromHost=toHost;
		}
		else {
			args.toHost=toHost;
			args.fromHost=createObject(definition.fromHost);
		}
		parameterMap.addMapping(definition.type, args);
		if (definition.isComplexType)
			parameterMap.addComplexType(definition.type);
	};
	for (var id in initialDefs)
		parameterMap.define(initialDefs[id]);
	return parameterMap;
};
OSF.OUtil.setNamespace("DispIdHost", OSF.DDA);
OSF.DDA.DispIdHost.Methods={
	InvokeMethod: "invokeMethod",
	AddEventHandler: "addEventHandler",
	RemoveEventHandler: "removeEventHandler",
	OpenDialog: "openDialog",
	CloseDialog: "closeDialog",
	MessageParent: "messageParent",
	SendMessage: "sendMessage"
};
OSF.DDA.DispIdHost.Delegates={
	ExecuteAsync: "executeAsync",
	RegisterEventAsync: "registerEventAsync",
	UnregisterEventAsync: "unregisterEventAsync",
	ParameterMap: "parameterMap",
	OpenDialog: "openDialog",
	CloseDialog: "closeDialog",
	MessageParent: "messageParent",
	SendMessage: "sendMessage"
};
OSF.DDA.DispIdHost.Facade=function OSF_DDA_DispIdHost_Facade(getDelegateMethods, parameterMap) {
	var dispIdMap={};
	var jsom=OSF.DDA.AsyncMethodNames;
	var did=OSF.DDA.MethodDispId;
	var methodMap={
		"GoToByIdAsync": did.dispidNavigateToMethod,
		"GetSelectedDataAsync": did.dispidGetSelectedDataMethod,
		"SetSelectedDataAsync": did.dispidSetSelectedDataMethod,
		"GetDocumentCopyChunkAsync": did.dispidGetDocumentCopyChunkMethod,
		"ReleaseDocumentCopyAsync": did.dispidReleaseDocumentCopyMethod,
		"GetDocumentCopyAsync": did.dispidGetDocumentCopyMethod,
		"AddFromSelectionAsync": did.dispidAddBindingFromSelectionMethod,
		"AddFromPromptAsync": did.dispidAddBindingFromPromptMethod,
		"AddFromNamedItemAsync": did.dispidAddBindingFromNamedItemMethod,
		"GetAllAsync": did.dispidGetAllBindingsMethod,
		"GetByIdAsync": did.dispidGetBindingMethod,
		"ReleaseByIdAsync": did.dispidReleaseBindingMethod,
		"GetDataAsync": did.dispidGetBindingDataMethod,
		"SetDataAsync": did.dispidSetBindingDataMethod,
		"AddRowsAsync": did.dispidAddRowsMethod,
		"AddColumnsAsync": did.dispidAddColumnsMethod,
		"DeleteAllDataValuesAsync": did.dispidClearAllRowsMethod,
		"RefreshAsync": did.dispidLoadSettingsMethod,
		"SaveAsync": did.dispidSaveSettingsMethod,
		"GetActiveViewAsync": did.dispidGetActiveViewMethod,
		"GetFilePropertiesAsync": did.dispidGetFilePropertiesMethod,
		"GetOfficeThemeAsync": did.dispidGetOfficeThemeMethod,
		"GetDocumentThemeAsync": did.dispidGetDocumentThemeMethod,
		"ClearFormatsAsync": did.dispidClearFormatsMethod,
		"SetTableOptionsAsync": did.dispidSetTableOptionsMethod,
		"SetFormatsAsync": did.dispidSetFormatsMethod,
		"GetAccessTokenAsync": did.dispidGetAccessTokenMethod,
		"ExecuteRichApiRequestAsync": did.dispidExecuteRichApiRequestMethod,
		"AppCommandInvocationCompletedAsync": did.dispidAppCommandInvocationCompletedMethod,
		"CloseContainerAsync": did.dispidCloseContainerMethod,
		"OpenBrowserWindow": did.dispidOpenBrowserWindow,
		"CreateDocumentAsync": did.dispidCreateDocumentMethod,
		"InsertFormAsync": did.dispidInsertFormMethod,
		"AddDataPartAsync": did.dispidAddDataPartMethod,
		"GetDataPartByIdAsync": did.dispidGetDataPartByIdMethod,
		"GetDataPartsByNameSpaceAsync": did.dispidGetDataPartsByNamespaceMethod,
		"GetPartXmlAsync": did.dispidGetDataPartXmlMethod,
		"GetPartNodesAsync": did.dispidGetDataPartNodesMethod,
		"DeleteDataPartAsync": did.dispidDeleteDataPartMethod,
		"GetNodeValueAsync": did.dispidGetDataNodeValueMethod,
		"GetNodeXmlAsync": did.dispidGetDataNodeXmlMethod,
		"GetRelativeNodesAsync": did.dispidGetDataNodesMethod,
		"SetNodeValueAsync": did.dispidSetDataNodeValueMethod,
		"SetNodeXmlAsync": did.dispidSetDataNodeXmlMethod,
		"AddDataPartNamespaceAsync": did.dispidAddDataNamespaceMethod,
		"GetDataPartNamespaceAsync": did.dispidGetDataUriByPrefixMethod,
		"GetDataPartPrefixAsync": did.dispidGetDataPrefixByUriMethod,
		"GetNodeTextAsync": did.dispidGetDataNodeTextMethod,
		"SetNodeTextAsync": did.dispidSetDataNodeTextMethod,
		"GetSelectedTask": did.dispidGetSelectedTaskMethod,
		"GetTask": did.dispidGetTaskMethod,
		"GetWSSUrl": did.dispidGetWSSUrlMethod,
		"GetTaskField": did.dispidGetTaskFieldMethod,
		"GetSelectedResource": did.dispidGetSelectedResourceMethod,
		"GetResourceField": did.dispidGetResourceFieldMethod,
		"GetProjectField": did.dispidGetProjectFieldMethod,
		"GetSelectedView": did.dispidGetSelectedViewMethod,
		"GetTaskByIndex": did.dispidGetTaskByIndexMethod,
		"GetResourceByIndex": did.dispidGetResourceByIndexMethod,
		"SetTaskField": did.dispidSetTaskFieldMethod,
		"SetResourceField": did.dispidSetResourceFieldMethod,
		"GetMaxTaskIndex": did.dispidGetMaxTaskIndexMethod,
		"GetMaxResourceIndex": did.dispidGetMaxResourceIndexMethod,
		"CreateTask": did.dispidCreateTaskMethod
	};
	for (var method in methodMap) {
		if (jsom[method]) {
			dispIdMap[jsom[method].id]=methodMap[method];
		}
	}
	jsom=OSF.DDA.SyncMethodNames;
	did=OSF.DDA.MethodDispId;
	var syncMethodMap={
		"MessageParent": did.dispidMessageParentMethod,
		"SendMessage": did.dispidSendMessageMethod
	};
	for (var method in syncMethodMap) {
		if (jsom[method]) {
			dispIdMap[jsom[method].id]=syncMethodMap[method];
		}
	}
	jsom=Microsoft.Office.WebExtension.EventType;
	did=OSF.DDA.EventDispId;
	var eventMap={
		"SettingsChanged": did.dispidSettingsChangedEvent,
		"DocumentSelectionChanged": did.dispidDocumentSelectionChangedEvent,
		"BindingSelectionChanged": did.dispidBindingSelectionChangedEvent,
		"BindingDataChanged": did.dispidBindingDataChangedEvent,
		"ActiveViewChanged": did.dispidActiveViewChangedEvent,
		"OfficeThemeChanged": did.dispidOfficeThemeChangedEvent,
		"DocumentThemeChanged": did.dispidDocumentThemeChangedEvent,
		"AppCommandInvoked": did.dispidAppCommandInvokedEvent,
		"DialogMessageReceived": did.dispidDialogMessageReceivedEvent,
		"DialogParentMessageReceived": did.dispidDialogParentMessageReceivedEvent,
		"ObjectDeleted": did.dispidObjectDeletedEvent,
		"ObjectSelectionChanged": did.dispidObjectSelectionChangedEvent,
		"ObjectDataChanged": did.dispidObjectDataChangedEvent,
		"ContentControlAdded": did.dispidContentControlAddedEvent,
		"RichApiMessage": did.dispidRichApiMessageEvent,
		"ItemChanged": did.dispidOlkItemSelectedChangedEvent,
		"RecipientsChanged": did.dispidOlkRecipientsChangedEvent,
		"AppointmentTimeChanged": did.dispidOlkAppointmentTimeChangedEvent,
		"RecurrenceChanged": did.dispidOlkRecurrenceChangedEvent,
		"TaskSelectionChanged": did.dispidTaskSelectionChangedEvent,
		"ResourceSelectionChanged": did.dispidResourceSelectionChangedEvent,
		"ViewSelectionChanged": did.dispidViewSelectionChangedEvent,
		"DataNodeInserted": did.dispidDataNodeAddedEvent,
		"DataNodeReplaced": did.dispidDataNodeReplacedEvent,
		"DataNodeDeleted": did.dispidDataNodeDeletedEvent
	};
	for (var event in eventMap) {
		if (jsom[event]) {
			dispIdMap[jsom[event]]=eventMap[event];
		}
	}
	function IsObjectEvent(dispId) {
		return (dispId==OSF.DDA.EventDispId.dispidObjectDeletedEvent ||
			dispId==OSF.DDA.EventDispId.dispidObjectSelectionChangedEvent ||
			dispId==OSF.DDA.EventDispId.dispidObjectDataChangedEvent ||
			dispId==OSF.DDA.EventDispId.dispidContentControlAddedEvent);
	}
	function onException(ex, asyncMethodCall, suppliedArgs, callArgs) {
		if (typeof ex=="number") {
			if (!callArgs) {
				callArgs=asyncMethodCall.getCallArgs(suppliedArgs);
			}
			OSF.DDA.issueAsyncResult(callArgs, ex, OSF.DDA.ErrorCodeManager.getErrorArgs(ex));
		}
		else {
			throw ex;
		}
	}
	;
	this[OSF.DDA.DispIdHost.Methods.InvokeMethod]=function OSF_DDA_DispIdHost_Facade$InvokeMethod(method, suppliedArguments, caller, privateState) {
		var callArgs;
		try {
			var methodName=method.id;
			var asyncMethodCall=OSF.DDA.AsyncMethodCalls[methodName];
			callArgs=asyncMethodCall.verifyAndExtractCall(suppliedArguments, caller, privateState);
			var dispId=dispIdMap[methodName];
			var delegate=getDelegateMethods(methodName);
			var richApiInExcelMethodSubstitution=null;
			if (window.Excel && window.Office.context.requirements.isSetSupported("RedirectV1Api")) {
				window.Excel._RedirectV1APIs=true;
			}
			if (window.Excel && window.Excel._RedirectV1APIs && (richApiInExcelMethodSubstitution=window.Excel._V1APIMap[methodName])) {
				var preprocessedCallArgs=OSF.OUtil.shallowCopy(callArgs);
				delete preprocessedCallArgs[Microsoft.Office.WebExtension.Parameters.AsyncContext];
				if (richApiInExcelMethodSubstitution.preprocess) {
					preprocessedCallArgs=richApiInExcelMethodSubstitution.preprocess(preprocessedCallArgs);
				}
				var ctx=new window.Excel.RequestContext();
				var result=richApiInExcelMethodSubstitution.call(ctx, preprocessedCallArgs);
				ctx.sync()
					.then(function () {
					var response=result.value;
					var status=response.status;
					delete response["status"];
					delete response["@odata.type"];
					if (richApiInExcelMethodSubstitution.postprocess) {
						response=richApiInExcelMethodSubstitution.postprocess(response, preprocessedCallArgs);
					}
					if (status !=0) {
						response=OSF.DDA.ErrorCodeManager.getErrorArgs(status);
					}
					OSF.DDA.issueAsyncResult(callArgs, status, response);
				})["catch"](function (error) {
					OSF.DDA.issueAsyncResult(callArgs, OSF.DDA.ErrorCodeManager.errorCodes.ooeFailure, null);
				});
			}
			else {
				var hostCallArgs;
				if (parameterMap.toHost) {
					hostCallArgs=parameterMap.toHost(dispId, callArgs);
				}
				else {
					hostCallArgs=callArgs;
				}
				var startTime=(new Date()).getTime();
				delegate[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]({
					"dispId": dispId,
					"hostCallArgs": hostCallArgs,
					"onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { },
					"onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { },
					"onComplete": function (status, hostResponseArgs) {
						var responseArgs;
						if (status==OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
							if (parameterMap.fromHost) {
								responseArgs=parameterMap.fromHost(dispId, hostResponseArgs);
							}
							else {
								responseArgs=hostResponseArgs;
							}
						}
						else {
							responseArgs=hostResponseArgs;
						}
						var payload=asyncMethodCall.processResponse(status, responseArgs, caller, callArgs);
						OSF.DDA.issueAsyncResult(callArgs, status, payload);
						if (OSF.AppTelemetry) {
							OSF.AppTelemetry.onMethodDone(dispId, hostCallArgs, Math.abs((new Date()).getTime() - startTime), status);
						}
					}
				});
			}
		}
		catch (ex) {
			onException(ex, asyncMethodCall, suppliedArguments, callArgs);
		}
	};
	this[OSF.DDA.DispIdHost.Methods.AddEventHandler]=function OSF_DDA_DispIdHost_Facade$AddEventHandler(suppliedArguments, eventDispatch, caller, isPopupWindow) {
		var callArgs;
		var eventType, handler;
		var isObjectEvent=false;
		function onEnsureRegistration(status) {
			if (status==OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
				var added=!isObjectEvent ? eventDispatch.addEventHandler(eventType, handler) :
					eventDispatch.addObjectEventHandler(eventType, callArgs[Microsoft.Office.WebExtension.Parameters.Id], handler);
				if (!added) {
					status=OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerAdditionFailed;
				}
			}
			var error;
			if (status !=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
				error=OSF.DDA.ErrorCodeManager.getErrorArgs(status);
			}
			OSF.DDA.issueAsyncResult(callArgs, status, error);
		}
		try {
			var asyncMethodCall=OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.AddHandlerAsync.id];
			callArgs=asyncMethodCall.verifyAndExtractCall(suppliedArguments, caller, eventDispatch);
			eventType=callArgs[Microsoft.Office.WebExtension.Parameters.EventType];
			handler=callArgs[Microsoft.Office.WebExtension.Parameters.Handler];
			if (isPopupWindow) {
				onEnsureRegistration(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess);
				return;
			}
			var dispId=dispIdMap[eventType];
			isObjectEvent=IsObjectEvent(dispId);
			var targetId=(isObjectEvent ? callArgs[Microsoft.Office.WebExtension.Parameters.Id] : (caller.id || ""));
			var count=isObjectEvent ? eventDispatch.getObjectEventHandlerCount(eventType, targetId) : eventDispatch.getEventHandlerCount(eventType);
			if (count==0) {
				var invoker=getDelegateMethods(eventType)[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync];
				invoker({
					"eventType": eventType,
					"dispId": dispId,
					"targetId": targetId,
					"onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall); },
					"onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse); },
					"onComplete": onEnsureRegistration,
					"onEvent": function handleEvent(hostArgs) {
						var args=parameterMap.fromHost(dispId, hostArgs);
						if (!isObjectEvent)
							eventDispatch.fireEvent(OSF.DDA.OMFactory.manufactureEventArgs(eventType, caller, args));
						else
							eventDispatch.fireObjectEvent(targetId, OSF.DDA.OMFactory.manufactureEventArgs(eventType, targetId, args));
					}
				});
			}
			else {
				onEnsureRegistration(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess);
			}
		}
		catch (ex) {
			onException(ex, asyncMethodCall, suppliedArguments, callArgs);
		}
	};
	this[OSF.DDA.DispIdHost.Methods.RemoveEventHandler]=function OSF_DDA_DispIdHost_Facade$RemoveEventHandler(suppliedArguments, eventDispatch, caller) {
		var callArgs;
		var eventType, handler;
		var isObjectEvent=false;
		function onEnsureRegistration(status) {
			var error;
			if (status !=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
				error=OSF.DDA.ErrorCodeManager.getErrorArgs(status);
			}
			OSF.DDA.issueAsyncResult(callArgs, status, error);
		}
		try {
			var asyncMethodCall=OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.RemoveHandlerAsync.id];
			callArgs=asyncMethodCall.verifyAndExtractCall(suppliedArguments, caller, eventDispatch);
			eventType=callArgs[Microsoft.Office.WebExtension.Parameters.EventType];
			handler=callArgs[Microsoft.Office.WebExtension.Parameters.Handler];
			var dispId=dispIdMap[eventType];
			isObjectEvent=IsObjectEvent(dispId);
			var targetId=(isObjectEvent ? callArgs[Microsoft.Office.WebExtension.Parameters.Id] : (caller.id || ""));
			var status, removeSuccess;
			if (handler===null) {
				removeSuccess=isObjectEvent ? eventDispatch.clearObjectEventHandlers(eventType, targetId) : eventDispatch.clearEventHandlers(eventType);
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess;
			}
			else {
				removeSuccess=isObjectEvent ? eventDispatch.removeObjectEventHandler(eventType, targetId, handler) : eventDispatch.removeEventHandler(eventType, handler);
				status=removeSuccess ? OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess : OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerNotExist;
			}
			var count=isObjectEvent ? eventDispatch.getObjectEventHandlerCount(eventType, targetId) : eventDispatch.getEventHandlerCount(eventType);
			if (removeSuccess && count==0) {
				var invoker=getDelegateMethods(eventType)[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync];
				invoker({
					"eventType": eventType,
					"dispId": dispId,
					"targetId": targetId,
					"onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall); },
					"onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse); },
					"onComplete": onEnsureRegistration
				});
			}
			else {
				onEnsureRegistration(status);
			}
		}
		catch (ex) {
			onException(ex, asyncMethodCall, suppliedArguments, callArgs);
		}
	};
	this[OSF.DDA.DispIdHost.Methods.OpenDialog]=function OSF_DDA_DispIdHost_Facade$OpenDialog(suppliedArguments, eventDispatch, caller) {
		var callArgs;
		var targetId;
		var dialogMessageEvent=Microsoft.Office.WebExtension.EventType.DialogMessageReceived;
		var dialogOtherEvent=Microsoft.Office.WebExtension.EventType.DialogEventReceived;
		function onEnsureRegistration(status) {
			var payload;
			if (status !=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
				payload=OSF.DDA.ErrorCodeManager.getErrorArgs(status);
			}
			else {
				var onSucceedArgs={};
				onSucceedArgs[Microsoft.Office.WebExtension.Parameters.Id]=targetId;
				onSucceedArgs[Microsoft.Office.WebExtension.Parameters.Data]=eventDispatch;
				var payload=asyncMethodCall.processResponse(status, onSucceedArgs, caller, callArgs);
				OSF.DialogShownStatus.hasDialogShown=true;
				eventDispatch.clearEventHandlers(dialogMessageEvent);
				eventDispatch.clearEventHandlers(dialogOtherEvent);
			}
			OSF.DDA.issueAsyncResult(callArgs, status, payload);
		}
		try {
			if (dialogMessageEvent==undefined || dialogOtherEvent==undefined) {
				onEnsureRegistration(OSF.DDA.ErrorCodeManager.ooeOperationNotSupported);
			}
			if (OSF.DDA.AsyncMethodNames.DisplayDialogAsync==null) {
				onEnsureRegistration(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);
				return;
			}
			var asyncMethodCall=OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.DisplayDialogAsync.id];
			callArgs=asyncMethodCall.verifyAndExtractCall(suppliedArguments, caller, eventDispatch);
			var dispId=dispIdMap[dialogMessageEvent];
			var delegateMethods=getDelegateMethods(dialogMessageEvent);
			var invoker=delegateMethods[OSF.DDA.DispIdHost.Delegates.OpenDialog] !=undefined ?
				delegateMethods[OSF.DDA.DispIdHost.Delegates.OpenDialog] :
				delegateMethods[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync];
			targetId=JSON.stringify(callArgs);
			if (!OSF.DialogShownStatus.hasDialogShown) {
				eventDispatch.clearQueuedEvent(dialogMessageEvent);
				eventDispatch.clearQueuedEvent(dialogOtherEvent);
				eventDispatch.clearQueuedEvent(Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived);
			}
			invoker({
				"eventType": dialogMessageEvent,
				"dispId": dispId,
				"targetId": targetId,
				"onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall); },
				"onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse); },
				"onComplete": onEnsureRegistration,
				"onEvent": function handleEvent(hostArgs) {
					var args=parameterMap.fromHost(dispId, hostArgs);
					var event=OSF.DDA.OMFactory.manufactureEventArgs(dialogMessageEvent, caller, args);
					if (event.type==dialogOtherEvent) {
						var payload=OSF.DDA.ErrorCodeManager.getErrorArgs(event.error);
						var errorArgs={};
						errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Code]=status || OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
						errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name]=payload.name || payload;
						errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message]=payload.message || payload;
						event.error=new OSF.DDA.Error(errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name], errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message], errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Code]);
					}
					eventDispatch.fireOrQueueEvent(event);
					if (args[OSF.DDA.PropertyDescriptors.MessageType]==OSF.DialogMessageType.DialogClosed) {
						eventDispatch.clearEventHandlers(dialogMessageEvent);
						eventDispatch.clearEventHandlers(dialogOtherEvent);
						eventDispatch.clearEventHandlers(Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived);
						OSF.DialogShownStatus.hasDialogShown=false;
					}
				}
			});
		}
		catch (ex) {
			onException(ex, asyncMethodCall, suppliedArguments, callArgs);
		}
	};
	this[OSF.DDA.DispIdHost.Methods.CloseDialog]=function OSF_DDA_DispIdHost_Facade$CloseDialog(suppliedArguments, targetId, eventDispatch, caller) {
		var callArgs;
		var dialogMessageEvent, dialogOtherEvent;
		var closeStatus=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess;
		function closeCallback(status) {
			closeStatus=status;
			OSF.DialogShownStatus.hasDialogShown=false;
		}
		try {
			var asyncMethodCall=OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.CloseAsync.id];
			callArgs=asyncMethodCall.verifyAndExtractCall(suppliedArguments, caller, eventDispatch);
			dialogMessageEvent=Microsoft.Office.WebExtension.EventType.DialogMessageReceived;
			dialogOtherEvent=Microsoft.Office.WebExtension.EventType.DialogEventReceived;
			eventDispatch.clearEventHandlers(dialogMessageEvent);
			eventDispatch.clearEventHandlers(dialogOtherEvent);
			var dispId=dispIdMap[dialogMessageEvent];
			var delegateMethods=getDelegateMethods(dialogMessageEvent);
			var invoker=delegateMethods[OSF.DDA.DispIdHost.Delegates.CloseDialog] !=undefined ?
				delegateMethods[OSF.DDA.DispIdHost.Delegates.CloseDialog] :
				delegateMethods[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync];
			invoker({
				"eventType": dialogMessageEvent,
				"dispId": dispId,
				"targetId": targetId,
				"onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall); },
				"onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse); },
				"onComplete": closeCallback
			});
		}
		catch (ex) {
			onException(ex, asyncMethodCall, suppliedArguments, callArgs);
		}
		if (closeStatus !=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
			throw OSF.OUtil.formatString(Strings.OfficeOM.L_FunctionCallFailed, OSF.DDA.AsyncMethodNames.CloseAsync.displayName, closeStatus);
		}
	};
	this[OSF.DDA.DispIdHost.Methods.MessageParent]=function OSF_DDA_DispIdHost_Facade$MessageParent(suppliedArguments, caller) {
		var stateInfo={};
		var syncMethodCall=OSF.DDA.SyncMethodCalls[OSF.DDA.SyncMethodNames.MessageParent.id];
		var callArgs=syncMethodCall.verifyAndExtractCall(suppliedArguments, caller, stateInfo);
		var delegate=getDelegateMethods(OSF.DDA.SyncMethodNames.MessageParent.id);
		var invoker=delegate[OSF.DDA.DispIdHost.Delegates.MessageParent];
		var dispId=dispIdMap[OSF.DDA.SyncMethodNames.MessageParent.id];
		return invoker({
			"dispId": dispId,
			"hostCallArgs": callArgs,
			"onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall); },
			"onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse); }
		});
	};
	this[OSF.DDA.DispIdHost.Methods.SendMessage]=function OSF_DDA_DispIdHost_Facade$SendMessage(suppliedArguments, eventDispatch, caller) {
		var stateInfo={};
		var syncMethodCall=OSF.DDA.SyncMethodCalls[OSF.DDA.SyncMethodNames.SendMessage.id];
		var callArgs=syncMethodCall.verifyAndExtractCall(suppliedArguments, caller, stateInfo);
		var delegate=getDelegateMethods(OSF.DDA.SyncMethodNames.SendMessage.id);
		var invoker=delegate[OSF.DDA.DispIdHost.Delegates.SendMessage];
		var dispId=dispIdMap[OSF.DDA.SyncMethodNames.SendMessage.id];
		return invoker({
			"dispId": dispId,
			"hostCallArgs": callArgs,
			"onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall); },
			"onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse); }
		});
	};
};
OSF.DDA.DispIdHost.addAsyncMethods=function OSF_DDA_DispIdHost$AddAsyncMethods(target, asyncMethodNames, privateState) {
	for (var entry in asyncMethodNames) {
		var method=asyncMethodNames[entry];
		var name=method.displayName;
		if (!target[name]) {
			OSF.OUtil.defineEnumerableProperty(target, name, {
				value: (function (asyncMethod) {
					return function () {
						var invokeMethod=OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.InvokeMethod];
						invokeMethod(asyncMethod, arguments, target, privateState);
					};
				})(method)
			});
		}
	}
};
OSF.DDA.DispIdHost.addEventSupport=function OSF_DDA_DispIdHost$AddEventSupport(target, eventDispatch, isPopupWindow) {
	var add=OSF.DDA.AsyncMethodNames.AddHandlerAsync.displayName;
	var remove=OSF.DDA.AsyncMethodNames.RemoveHandlerAsync.displayName;
	if (!target[add]) {
		OSF.OUtil.defineEnumerableProperty(target, add, {
			value: function () {
				var addEventHandler=OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.AddEventHandler];
				addEventHandler(arguments, eventDispatch, target, isPopupWindow);
			}
		});
	}
	if (!target[remove]) {
		OSF.OUtil.defineEnumerableProperty(target, remove, {
			value: function () {
				var removeEventHandler=OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.RemoveEventHandler];
				removeEventHandler(arguments, eventDispatch, target);
			}
		});
	}
};
var OfficeExt;
(function (OfficeExt) {
	var MsAjaxTypeHelper=(function () {
		function MsAjaxTypeHelper() {
		}
		MsAjaxTypeHelper.isInstanceOfType=function (type, instance) {
			if (typeof (instance)==="undefined" || instance===null)
				return false;
			if (instance instanceof type)
				return true;
			var instanceType=instance.constructor;
			if (!instanceType || (typeof (instanceType) !=="function") || !instanceType.__typeName || instanceType.__typeName==='Object') {
				instanceType=Object;
			}
			return !!(instanceType===type) ||
				(instanceType.__typeName && type.__typeName && instanceType.__typeName===type.__typeName);
		};
		return MsAjaxTypeHelper;
	})();
	OfficeExt.MsAjaxTypeHelper=MsAjaxTypeHelper;
	var MsAjaxError=(function () {
		function MsAjaxError() {
		}
		MsAjaxError.create=function (message, errorInfo) {
			var err=new Error(message);
			err.message=message;
			if (errorInfo) {
				for (var v in errorInfo) {
					err[v]=errorInfo[v];
				}
			}
			err.popStackFrame();
			return err;
		};
		MsAjaxError.parameterCount=function (message) {
			var displayMessage="Sys.ParameterCountException: "+(message ? message : "Parameter count mismatch.");
			var err=MsAjaxError.create(displayMessage, { name: 'Sys.ParameterCountException' });
			err.popStackFrame();
			return err;
		};
		MsAjaxError.argument=function (paramName, message) {
			var displayMessage="Sys.ArgumentException: "+(message ? message : "Value does not fall within the expected range.");
			if (paramName) {
				displayMessage+="\n"+MsAjaxString.format("Parameter name: {0}", paramName);
			}
			var err=MsAjaxError.create(displayMessage, { name: "Sys.ArgumentException", paramName: paramName });
			err.popStackFrame();
			return err;
		};
		MsAjaxError.argumentNull=function (paramName, message) {
			var displayMessage="Sys.ArgumentNullException: "+(message ? message : "Value cannot be null.");
			if (paramName) {
				displayMessage+="\n"+MsAjaxString.format("Parameter name: {0}", paramName);
			}
			var err=MsAjaxError.create(displayMessage, { name: "Sys.ArgumentNullException", paramName: paramName });
			err.popStackFrame();
			return err;
		};
		MsAjaxError.argumentOutOfRange=function (paramName, actualValue, message) {
			var displayMessage="Sys.ArgumentOutOfRangeException: "+(message ? message : "Specified argument was out of the range of valid values.");
			if (paramName) {
				displayMessage+="\n"+MsAjaxString.format("Parameter name: {0}", paramName);
			}
			if (typeof (actualValue) !=="undefined" && actualValue !==null) {
				displayMessage+="\n"+MsAjaxString.format("Actual value was {0}.", actualValue);
			}
			var err=MsAjaxError.create(displayMessage, {
				name: "Sys.ArgumentOutOfRangeException",
				paramName: paramName,
				actualValue: actualValue
			});
			err.popStackFrame();
			return err;
		};
		MsAjaxError.argumentType=function (paramName, actualType, expectedType, message) {
			var displayMessage="Sys.ArgumentTypeException: ";
			if (message) {
				displayMessage+=message;
			}
			else if (actualType && expectedType) {
				displayMessage+=MsAjaxString.format("Object of type '{0}' cannot be converted to type '{1}'.", actualType.getName ? actualType.getName() : actualType, expectedType.getName ? expectedType.getName() : expectedType);
			}
			else {
				displayMessage+="Object cannot be converted to the required type.";
			}
			if (paramName) {
				displayMessage+="\n"+MsAjaxString.format("Parameter name: {0}", paramName);
			}
			var err=MsAjaxError.create(displayMessage, {
				name: "Sys.ArgumentTypeException",
				paramName: paramName,
				actualType: actualType,
				expectedType: expectedType
			});
			err.popStackFrame();
			return err;
		};
		MsAjaxError.argumentUndefined=function (paramName, message) {
			var displayMessage="Sys.ArgumentUndefinedException: "+(message ? message : "Value cannot be undefined.");
			if (paramName) {
				displayMessage+="\n"+MsAjaxString.format("Parameter name: {0}", paramName);
			}
			var err=MsAjaxError.create(displayMessage, { name: "Sys.ArgumentUndefinedException", paramName: paramName });
			err.popStackFrame();
			return err;
		};
		MsAjaxError.invalidOperation=function (message) {
			var displayMessage="Sys.InvalidOperationException: "+(message ? message : "Operation is not valid due to the current state of the object.");
			var err=MsAjaxError.create(displayMessage, { name: 'Sys.InvalidOperationException' });
			err.popStackFrame();
			return err;
		};
		return MsAjaxError;
	})();
	OfficeExt.MsAjaxError=MsAjaxError;
	var MsAjaxString=(function () {
		function MsAjaxString() {
		}
		MsAjaxString.format=function (format) {
			var args=[];
			for (var _i=1; _i < arguments.length; _i++) {
				args[_i - 1]=arguments[_i];
			}
			var source=format;
			return source.replace(/{(\d+)}/gm, function (match, number) {
				var index=parseInt(number, 10);
				return args[index]===undefined ? '{'+number+'}' : args[index];
			});
		};
		MsAjaxString.startsWith=function (str, prefix) {
			return (str.substr(0, prefix.length)===prefix);
		};
		return MsAjaxString;
	})();
	OfficeExt.MsAjaxString=MsAjaxString;
	var MsAjaxDebug=(function () {
		function MsAjaxDebug() {
		}
		MsAjaxDebug.trace=function (text) {
			if (typeof Debug !=="undefined" && Debug.writeln)
				Debug.writeln(text);
			if (window.console && window.console.log)
				window.console.log(text);
			if (window.opera && window.opera.postError)
				window.opera.postError(text);
			if (window.debugService && window.debugService.trace)
				window.debugService.trace(text);
			var a=document.getElementById("TraceConsole");
			if (a && a.tagName.toUpperCase()==="TEXTAREA") {
				a.innerHTML+=text+"\n";
			}
		};
		return MsAjaxDebug;
	})();
	OfficeExt.MsAjaxDebug=MsAjaxDebug;
	if (!OsfMsAjaxFactory.isMsAjaxLoaded()) {
		var registerTypeInternal=function registerTypeInternal(type, name, isClass) {
			if (type.__typeName===undefined) {
				type.__typeName=name;
			}
			if (type.__class===undefined) {
				type.__class=isClass;
			}
		};
		registerTypeInternal(Function, "Function", true);
		registerTypeInternal(Error, "Error", true);
		registerTypeInternal(Object, "Object", true);
		registerTypeInternal(String, "String", true);
		registerTypeInternal(Boolean, "Boolean", true);
		registerTypeInternal(Date, "Date", true);
		registerTypeInternal(Number, "Number", true);
		registerTypeInternal(RegExp, "RegExp", true);
		registerTypeInternal(Array, "Array", true);
		if (!Function.createCallback) {
			Function.createCallback=function Function$createCallback(method, context) {
				var e=Function._validateParams(arguments, [
					{ name: "method", type: Function },
					{ name: "context", mayBeNull: true }
				]);
				if (e)
					throw e;
				return function () {
					var l=arguments.length;
					if (l > 0) {
						var args=[];
						for (var i=0; i < l; i++) {
							args[i]=arguments[i];
						}
						args[l]=context;
						return method.apply(this, args);
					}
					return method.call(this, context);
				};
			};
		}
		if (!Function.createDelegate) {
			Function.createDelegate=function Function$createDelegate(instance, method) {
				var e=Function._validateParams(arguments, [
					{ name: "instance", mayBeNull: true },
					{ name: "method", type: Function }
				]);
				if (e)
					throw e;
				return function () {
					return method.apply(instance, arguments);
				};
			};
		}
		if (!Function._validateParams) {
			Function._validateParams=function (params, expectedParams, validateParameterCount) {
				var e, expectedLength=expectedParams.length;
				validateParameterCount=validateParameterCount || (typeof (validateParameterCount)==="undefined");
				e=Function._validateParameterCount(params, expectedParams, validateParameterCount);
				if (e) {
					e.popStackFrame();
					return e;
				}
				for (var i=0, l=params.length; i < l; i++) {
					var expectedParam=expectedParams[Math.min(i, expectedLength - 1)], paramName=expectedParam.name;
					if (expectedParam.parameterArray) {
						paramName+="["+(i - expectedLength+1)+"]";
					}
					else if (!validateParameterCount && (i >=expectedLength)) {
						break;
					}
					e=Function._validateParameter(params[i], expectedParam, paramName);
					if (e) {
						e.popStackFrame();
						return e;
					}
				}
				return null;
			};
		}
		if (!Function._validateParameterCount) {
			Function._validateParameterCount=function (params, expectedParams, validateParameterCount) {
				var i, error, expectedLen=expectedParams.length, actualLen=params.length;
				if (actualLen < expectedLen) {
					var minParams=expectedLen;
					for (i=0; i < expectedLen; i++) {
						var param=expectedParams[i];
						if (param.optional || param.parameterArray) {
							minParams--;
						}
					}
					if (actualLen < minParams) {
						error=true;
					}
				}
				else if (validateParameterCount && (actualLen > expectedLen)) {
					error=true;
					for (i=0; i < expectedLen; i++) {
						if (expectedParams[i].parameterArray) {
							error=false;
							break;
						}
					}
				}
				if (error) {
					var e=MsAjaxError.parameterCount();
					e.popStackFrame();
					return e;
				}
				return null;
			};
		}
		if (!Function._validateParameter) {
			Function._validateParameter=function (param, expectedParam, paramName) {
				var e, expectedType=expectedParam.type, expectedInteger=!!expectedParam.integer, expectedDomElement=!!expectedParam.domElement, mayBeNull=!!expectedParam.mayBeNull;
				e=Function._validateParameterType(param, expectedType, expectedInteger, expectedDomElement, mayBeNull, paramName);
				if (e) {
					e.popStackFrame();
					return e;
				}
				var expectedElementType=expectedParam.elementType, elementMayBeNull=!!expectedParam.elementMayBeNull;
				if (expectedType===Array && typeof (param) !=="undefined" && param !==null &&
					(expectedElementType || !elementMayBeNull)) {
					var expectedElementInteger=!!expectedParam.elementInteger, expectedElementDomElement=!!expectedParam.elementDomElement;
					for (var i=0; i < param.length; i++) {
						var elem=param[i];
						e=Function._validateParameterType(elem, expectedElementType, expectedElementInteger, expectedElementDomElement, elementMayBeNull, paramName+"["+i+"]");
						if (e) {
							e.popStackFrame();
							return e;
						}
					}
				}
				return null;
			};
		}
		if (!Function._validateParameterType) {
			Function._validateParameterType=function (param, expectedType, expectedInteger, expectedDomElement, mayBeNull, paramName) {
				var e, i;
				if (typeof (param)==="undefined") {
					if (mayBeNull) {
						return null;
					}
					else {
						e=OfficeExt.MsAjaxError.argumentUndefined(paramName);
						e.popStackFrame();
						return e;
					}
				}
				if (param===null) {
					if (mayBeNull) {
						return null;
					}
					else {
						e=OfficeExt.MsAjaxError.argumentNull(paramName);
						e.popStackFrame();
						return e;
					}
				}
				if (expectedType && !OfficeExt.MsAjaxTypeHelper.isInstanceOfType(expectedType, param)) {
					e=OfficeExt.MsAjaxError.argumentType(paramName, typeof (param), expectedType);
					e.popStackFrame();
					return e;
				}
				return null;
			};
		}
		if (!window.Type) {
			window.Type=Function;
		}
		if (!Type.registerNamespace) {
			Type.registerNamespace=function (ns) {
				var namespaceParts=ns.split('.');
				var currentNamespace=window;
				for (var i=0; i < namespaceParts.length; i++) {
					currentNamespace[namespaceParts[i]]=currentNamespace[namespaceParts[i]] || {};
					currentNamespace=currentNamespace[namespaceParts[i]];
				}
			};
		}
		if (!Type.prototype.registerClass) {
			Type.prototype.registerClass=function (cls) { cls={}; };
		}
		if (typeof (Sys)==="undefined") {
			Type.registerNamespace('Sys');
		}
		if (!Error.prototype.popStackFrame) {
			Error.prototype.popStackFrame=function () {
				if (arguments.length !==0)
					throw MsAjaxError.parameterCount();
				if (typeof (this.stack)==="undefined" || this.stack===null ||
					typeof (this.fileName)==="undefined" || this.fileName===null ||
					typeof (this.lineNumber)==="undefined" || this.lineNumber===null) {
					return;
				}
				var stackFrames=this.stack.split("\n");
				var currentFrame=stackFrames[0];
				var pattern=this.fileName+":"+this.lineNumber;
				while (typeof (currentFrame) !=="undefined" &&
					currentFrame !==null &&
					currentFrame.indexOf(pattern)===-1) {
					stackFrames.shift();
					currentFrame=stackFrames[0];
				}
				var nextFrame=stackFrames[1];
				if (typeof (nextFrame)==="undefined" || nextFrame===null) {
					return;
				}
				var nextFrameParts=nextFrame.match(/@(.*):(\d+)$/);
				if (typeof (nextFrameParts)==="undefined" || nextFrameParts===null) {
					return;
				}
				this.fileName=nextFrameParts[1];
				this.lineNumber=parseInt(nextFrameParts[2]);
				stackFrames.shift();
				this.stack=stackFrames.join("\n");
			};
		}
		OsfMsAjaxFactory.msAjaxError=MsAjaxError;
		OsfMsAjaxFactory.msAjaxString=MsAjaxString;
		OsfMsAjaxFactory.msAjaxDebug=MsAjaxDebug;
	}
})(OfficeExt || (OfficeExt={}));
OSF.OUtil.setNamespace("SafeArray", OSF.DDA);
OSF.DDA.SafeArray.Response={
	Status: 0,
	Payload: 1
};
OSF.DDA.SafeArray.UniqueArguments={
	Offset: "offset",
	Run: "run",
	BindingSpecificData: "bindingSpecificData",
	MergedCellGuid: "{66e7831f-81b2-42e2-823c-89e872d541b3}"
};
OSF.OUtil.setNamespace("Delegate", OSF.DDA.SafeArray);
OSF.DDA.SafeArray.Delegate._onException=function OSF_DDA_SafeArray_Delegate$OnException(ex, args) {
	var status;
	var statusNumber=ex.number;
	if (statusNumber) {
		switch (statusNumber) {
			case -2146828218:
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability;
				break;
			case -2147467259:
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeDialogAlreadyOpened;
				break;
			case -2146828283:
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidParam;
				break;
			case -2147209089:
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidParam;
				break;
			case -2146827850:
			default:
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
				break;
		}
	}
	if (args.onComplete) {
		args.onComplete(status || OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);
	}
};
OSF.DDA.SafeArray.Delegate._onExceptionSyncMethod=function OSF_DDA_SafeArray_Delegate$OnExceptionSyncMethod(ex, args) {
	var status;
	var number=ex.number;
	if (number) {
		switch (number) {
			case -2146828218:
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability;
				break;
			case -2146827850:
			default:
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
				break;
		}
	}
	return status || OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
};
OSF.DDA.SafeArray.Delegate.SpecialProcessor=function OSF_DDA_SafeArray_Delegate_SpecialProcessor() {
	function _2DVBArrayToJaggedArray(vbArr) {
		var ret;
		try {
			var rows=vbArr.ubound(1);
			var cols=vbArr.ubound(2);
			vbArr=vbArr.toArray();
			if (rows==1 && cols==1) {
				ret=[vbArr];
			}
			else {
				ret=[];
				for (var row=0; row < rows; row++) {
					var rowArr=[];
					for (var col=0; col < cols; col++) {
						var datum=vbArr[row * cols+col];
						if (datum !=OSF.DDA.SafeArray.UniqueArguments.MergedCellGuid) {
							rowArr.push(datum);
						}
					}
					if (rowArr.length > 0) {
						ret.push(rowArr);
					}
				}
			}
		}
		catch (ex) {
		}
		return ret;
	}
	var complexTypes=[];
	var dynamicTypes={};
	dynamicTypes[Microsoft.Office.WebExtension.Parameters.Data]=(function () {
		var tableRows=0;
		var tableHeaders=1;
		return {
			toHost: function OSF_DDA_SafeArray_Delegate_SpecialProcessor_Data$toHost(data) {
				if (OSF.DDA.TableDataProperties && typeof data !="string" && data[OSF.DDA.TableDataProperties.TableRows] !==undefined) {
					var tableData=[];
					tableData[tableRows]=data[OSF.DDA.TableDataProperties.TableRows];
					tableData[tableHeaders]=data[OSF.DDA.TableDataProperties.TableHeaders];
					data=tableData;
				}
				return data;
			},
			fromHost: function OSF_DDA_SafeArray_Delegate_SpecialProcessor_Data$fromHost(hostArgs) {
				var ret;
				if (hostArgs.toArray) {
					var dimensions=hostArgs.dimensions();
					if (dimensions===2) {
						ret=_2DVBArrayToJaggedArray(hostArgs);
					}
					else {
						var array=hostArgs.toArray();
						if (array.length===2 && ((array[0] !=null && array[0].toArray) || (array[1] !=null && array[1].toArray))) {
							ret={};
							ret[OSF.DDA.TableDataProperties.TableRows]=_2DVBArrayToJaggedArray(array[tableRows]);
							ret[OSF.DDA.TableDataProperties.TableHeaders]=_2DVBArrayToJaggedArray(array[tableHeaders]);
						}
						else {
							ret=array;
						}
					}
				}
				else {
					ret=hostArgs;
				}
				return ret;
			}
		};
	})();
	OSF.DDA.SafeArray.Delegate.SpecialProcessor.uber.constructor.call(this, complexTypes, dynamicTypes);
	this.unpack=function OSF_DDA_SafeArray_Delegate_SpecialProcessor$unpack(param, arg) {
		var value;
		if (this.isComplexType(param) || OSF.DDA.ListType.isListType(param)) {
			var toArraySupported=(arg || typeof arg==="unknown") && arg.toArray;
			value=toArraySupported ? arg.toArray() : arg || {};
		}
		else if (this.isDynamicType(param)) {
			value=dynamicTypes[param].fromHost(arg);
		}
		else {
			value=arg;
		}
		return value;
	};
};
OSF.OUtil.extend(OSF.DDA.SafeArray.Delegate.SpecialProcessor, OSF.DDA.SpecialProcessor);
OSF.DDA.SafeArray.Delegate.ParameterMap=OSF.DDA.getDecoratedParameterMap(new OSF.DDA.SafeArray.Delegate.SpecialProcessor(), [
	{
		type: Microsoft.Office.WebExtension.Parameters.ValueFormat,
		toHost: [
			{ name: Microsoft.Office.WebExtension.ValueFormat.Unformatted, value: 0 },
			{ name: Microsoft.Office.WebExtension.ValueFormat.Formatted, value: 1 }
		]
	},
	{
		type: Microsoft.Office.WebExtension.Parameters.FilterType,
		toHost: [
			{ name: Microsoft.Office.WebExtension.FilterType.All, value: 0 }
		]
	}
]);
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.PropertyDescriptors.AsyncResultStatus,
	fromHost: [
		{ name: Microsoft.Office.WebExtension.AsyncResultStatus.Succeeded, value: 0 },
		{ name: Microsoft.Office.WebExtension.AsyncResultStatus.Failed, value: 1 }
	]
});
OSF.DDA.SafeArray.Delegate.executeAsync=function OSF_DDA_SafeArray_Delegate$ExecuteAsync(args) {
	function toArray(args) {
		var arrArgs=args;
		if (OSF.OUtil.isArray(args)) {
			var len=arrArgs.length;
			for (var i=0; i < len; i++) {
				arrArgs[i]=toArray(arrArgs[i]);
			}
		}
		else if (OSF.OUtil.isDate(args)) {
			arrArgs=args.getVarDate();
		}
		else if (typeof args==="object" && !OSF.OUtil.isArray(args)) {
			arrArgs=[];
			for (var index in args) {
				if (!OSF.OUtil.isFunction(args[index])) {
					arrArgs[index]=toArray(args[index]);
				}
			}
		}
		return arrArgs;
	}
	function fromSafeArray(value) {
		var ret=value;
		if (value !=null && value.toArray) {
			var arrayResult=value.toArray();
			ret=new Array(arrayResult.length);
			for (var i=0; i < arrayResult.length; i++) {
				ret[i]=fromSafeArray(arrayResult[i]);
			}
		}
		return ret;
	}
	try {
		if (args.onCalling) {
			args.onCalling();
		}
		OSF.ClientHostController.execute(args.dispId, toArray(args.hostCallArgs), function OSF_DDA_SafeArrayFacade$Execute_OnResponse(hostResponseArgs, resultCode) {
			var result=hostResponseArgs.toArray();
			var status=result[OSF.DDA.SafeArray.Response.Status];
			if (status==OSF.DDA.ErrorCodeManager.errorCodes.ooeChunkResult) {
				var payload=result[OSF.DDA.SafeArray.Response.Payload];
				payload=fromSafeArray(payload);
				if (payload !=null) {
					if (!args._chunkResultData) {
						args._chunkResultData=new Array();
					}
					args._chunkResultData[payload[0]]=payload[1];
				}
				return false;
			}
			if (args.onReceiving) {
				args.onReceiving();
			}
			if (args.onComplete) {
				var payload;
				if (status==OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
					if (result.length > 2) {
						payload=[];
						for (var i=1; i < result.length; i++)
							payload[i - 1]=result[i];
					}
					else {
						payload=result[OSF.DDA.SafeArray.Response.Payload];
					}
					if (args._chunkResultData) {
						payload=fromSafeArray(payload);
						if (payload !=null) {
							var expectedChunkCount=payload[payload.length - 1];
							if (args._chunkResultData.length==expectedChunkCount) {
								payload[payload.length - 1]=args._chunkResultData;
							}
							else {
								status=OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
							}
						}
					}
				}
				else {
					payload=result[OSF.DDA.SafeArray.Response.Payload];
				}
				args.onComplete(status, payload);
			}
			return true;
		});
	}
	catch (ex) {
		OSF.DDA.SafeArray.Delegate._onException(ex, args);
	}
};
OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent=function OSF_DDA_SafeArrayDelegate$GetOnAfterRegisterEvent(register, args) {
	var startTime=(new Date()).getTime();
	return function OSF_DDA_SafeArrayDelegate$OnAfterRegisterEvent(hostResponseArgs) {
		if (args.onReceiving) {
			args.onReceiving();
		}
		var status=hostResponseArgs.toArray ? hostResponseArgs.toArray()[OSF.DDA.SafeArray.Response.Status] : hostResponseArgs;
		if (args.onComplete) {
			args.onComplete(status);
		}
		if (OSF.AppTelemetry) {
			OSF.AppTelemetry.onRegisterDone(register, args.dispId, Math.abs((new Date()).getTime() - startTime), status);
		}
	};
};
OSF.DDA.SafeArray.Delegate.registerEventAsync=function OSF_DDA_SafeArray_Delegate$RegisterEventAsync(args) {
	if (args.onCalling) {
		args.onCalling();
	}
	var callback=OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent(true, args);
	try {
		OSF.ClientHostController.registerEvent(args.dispId, args.targetId, function OSF_DDA_SafeArrayDelegate$RegisterEventAsync_OnEvent(eventDispId, payload) {
			if (args.onEvent) {
				args.onEvent(payload);
			}
			if (OSF.AppTelemetry) {
				OSF.AppTelemetry.onEventDone(args.dispId);
			}
		}, callback);
	}
	catch (ex) {
		OSF.DDA.SafeArray.Delegate._onException(ex, args);
	}
};
OSF.DDA.SafeArray.Delegate.unregisterEventAsync=function OSF_DDA_SafeArray_Delegate$UnregisterEventAsync(args) {
	if (args.onCalling) {
		args.onCalling();
	}
	var callback=OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent(false, args);
	try {
		OSF.ClientHostController.unregisterEvent(args.dispId, args.targetId, callback);
	}
	catch (ex) {
		OSF.DDA.SafeArray.Delegate._onException(ex, args);
	}
};
OSF.ClientMode={
	ReadWrite: 0,
	ReadOnly: 1
};
OSF.DDA.RichInitializationReason={
	1: Microsoft.Office.WebExtension.InitializationReason.Inserted,
	2: Microsoft.Office.WebExtension.InitializationReason.DocumentOpened
};
OSF.InitializationHelper=function OSF_InitializationHelper(hostInfo, webAppState, context, settings, hostFacade) {
	this._hostInfo=hostInfo;
	this._webAppState=webAppState;
	this._context=context;
	this._settings=settings;
	this._hostFacade=hostFacade;
	this._initializeSettings=this.initializeSettings;
};
OSF.InitializationHelper.prototype.deserializeSettings=function OSF_InitializationHelper$deserializeSettings(serializedSettings, refreshSupported) {
	var settings;
	var osfSessionStorage=OSF.OUtil.getSessionStorage();
	if (osfSessionStorage) {
		var storageSettings=osfSessionStorage.getItem(OSF._OfficeAppFactory.getCachedSessionSettingsKey());
		if (storageSettings) {
			serializedSettings=JSON.parse(storageSettings);
		}
		else {
			storageSettings=JSON.stringify(serializedSettings);
			osfSessionStorage.setItem(OSF._OfficeAppFactory.getCachedSessionSettingsKey(), storageSettings);
		}
	}
	var deserializedSettings=OSF.DDA.SettingsManager.deserializeSettings(serializedSettings);
	if (refreshSupported) {
		settings=new OSF.DDA.RefreshableSettings(deserializedSettings);
	}
	else {
		settings=new OSF.DDA.Settings(deserializedSettings);
	}
	return settings;
};
OSF.InitializationHelper.prototype.saveAndSetDialogInfo=function OSF_InitializationHelper$saveAndSetDialogInfo(hostInfoValue) {
};
OSF.InitializationHelper.prototype.setAgaveHostCommunication=function OSF_InitializationHelper$setAgaveHostCommunication() {
};
OSF.InitializationHelper.prototype.prepareRightBeforeWebExtensionInitialize=function OSF_InitializationHelper$prepareRightBeforeWebExtensionInitialize(appContext) {
	this.prepareApiSurface(appContext);
	Microsoft.Office.WebExtension.initialize(this.getInitializationReason(appContext));
};
OSF.InitializationHelper.prototype.prepareApiSurface=function OSF_InitializationHelper$prepareApiSurfaceAndInitialize(appContext) {
	var license=new OSF.DDA.License(appContext.get_eToken());
	var getOfficeThemeHandler=(OSF.DDA.OfficeTheme && OSF.DDA.OfficeTheme.getOfficeTheme) ? OSF.DDA.OfficeTheme.getOfficeTheme : null;
	if (appContext.get_isDialog()) {
		if (OSF.DDA.UI.ChildUI) {
			appContext.ui=new OSF.DDA.UI.ChildUI();
		}
	}
	else {
		if (OSF.DDA.UI.ParentUI) {
			appContext.ui=new OSF.DDA.UI.ParentUI();
			if (OfficeExt.Container) {
				OSF.DDA.DispIdHost.addAsyncMethods(appContext.ui, [OSF.DDA.AsyncMethodNames.CloseContainerAsync]);
			}
		}
	}
	if (OSF.DDA.OpenBrowser) {
		OSF.DDA.DispIdHost.addAsyncMethods(appContext.ui, [OSF.DDA.AsyncMethodNames.OpenBrowserWindow]);
	}
	if (OSF.DDA.Auth) {
		appContext.auth=new OSF.DDA.Auth();
		OSF.DDA.DispIdHost.addAsyncMethods(appContext.auth, [OSF.DDA.AsyncMethodNames.GetAccessTokenAsync]);
	}
	OSF._OfficeAppFactory.setContext(new OSF.DDA.Context(appContext, appContext.doc, license, null, getOfficeThemeHandler));
	var getDelegateMethods, parameterMap;
	getDelegateMethods=OSF.DDA.DispIdHost.getClientDelegateMethods;
	parameterMap=OSF.DDA.SafeArray.Delegate.ParameterMap;
	OSF._OfficeAppFactory.setHostFacade(new OSF.DDA.DispIdHost.Facade(getDelegateMethods, parameterMap));
};
OSF.InitializationHelper.prototype.getInitializationReason=function (appContext) { return OSF.DDA.RichInitializationReason[appContext.get_reason()]; };
OSF.DDA.DispIdHost.getClientDelegateMethods=function (actionId) {
	var delegateMethods={};
	delegateMethods[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]=OSF.DDA.SafeArray.Delegate.executeAsync;
	delegateMethods[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync]=OSF.DDA.SafeArray.Delegate.registerEventAsync;
	delegateMethods[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync]=OSF.DDA.SafeArray.Delegate.unregisterEventAsync;
	delegateMethods[OSF.DDA.DispIdHost.Delegates.OpenDialog]=OSF.DDA.SafeArray.Delegate.openDialog;
	delegateMethods[OSF.DDA.DispIdHost.Delegates.CloseDialog]=OSF.DDA.SafeArray.Delegate.closeDialog;
	delegateMethods[OSF.DDA.DispIdHost.Delegates.MessageParent]=OSF.DDA.SafeArray.Delegate.messageParent;
	delegateMethods[OSF.DDA.DispIdHost.Delegates.SendMessage]=OSF.DDA.SafeArray.Delegate.sendMessage;
	if (OSF.DDA.AsyncMethodNames.RefreshAsync && actionId==OSF.DDA.AsyncMethodNames.RefreshAsync.id) {
		var readSerializedSettings=function (hostCallArgs, onCalling, onReceiving) {
			return OSF.DDA.ClientSettingsManager.read(onCalling, onReceiving);
		};
		delegateMethods[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]=OSF.DDA.ClientSettingsManager.getSettingsExecuteMethod(readSerializedSettings);
	}
	if (OSF.DDA.AsyncMethodNames.SaveAsync && actionId==OSF.DDA.AsyncMethodNames.SaveAsync.id) {
		var writeSerializedSettings=function (hostCallArgs, onCalling, onReceiving) {
			return OSF.DDA.ClientSettingsManager.write(hostCallArgs[OSF.DDA.SettingsManager.SerializedSettings], hostCallArgs[Microsoft.Office.WebExtension.Parameters.OverwriteIfStale], onCalling, onReceiving);
		};
		delegateMethods[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]=OSF.DDA.ClientSettingsManager.getSettingsExecuteMethod(writeSerializedSettings);
	}
	return delegateMethods;
};
var OSF=OSF || {};
var OSFWebkit;
(function (OSFWebkit) {
	var WebkitSafeArray=(function () {
		function WebkitSafeArray(data) {
			this.data=data;
			this.safeArrayFlag=this.isSafeArray(data);
		}
		WebkitSafeArray.prototype.dimensions=function () {
			var dimensions=0;
			if (this.safeArrayFlag) {
				dimensions=this.data[0][0];
			}
			else if (this.isArray()) {
				dimensions=2;
			}
			return dimensions;
		};
		WebkitSafeArray.prototype.getItem=function () {
			var array=[];
			var element=null;
			if (this.safeArrayFlag) {
				array=this.toArray();
			}
			else {
				array=this.data;
			}
			element=array;
			for (var i=0; i < arguments.length; i++) {
				element=element[arguments[i]];
			}
			return element;
		};
		WebkitSafeArray.prototype.lbound=function (dimension) {
			return 0;
		};
		WebkitSafeArray.prototype.ubound=function (dimension) {
			var ubound=0;
			if (this.safeArrayFlag) {
				ubound=this.data[0][dimension];
			}
			else if (this.isArray()) {
				if (dimension==1) {
					return this.data.length;
				}
				else if (dimension==2) {
					if (OSF.OUtil.isArray(this.data[0])) {
						return this.data[0].length;
					}
					else if (this.data[0] !=null) {
						return 1;
					}
				}
			}
			return ubound;
		};
		WebkitSafeArray.prototype.toArray=function () {
			if (this.isArray()==false) {
				return this.data;
			}
			var arr=[];
			var startingIndex=this.safeArrayFlag ? 1 : 0;
			for (var i=startingIndex; i < this.data.length; i++) {
				var element=this.data[i];
				if (this.isSafeArray(element)) {
					arr.push(new WebkitSafeArray(element));
				}
				else {
					arr.push(element);
				}
			}
			return arr;
		};
		WebkitSafeArray.prototype.isArray=function () {
			return OSF.OUtil.isArray(this.data);
		};
		WebkitSafeArray.prototype.isSafeArray=function (obj) {
			var isSafeArray=false;
			if (OSF.OUtil.isArray(obj) && OSF.OUtil.isArray(obj[0])) {
				var bounds=obj[0];
				var dimensions=bounds[0];
				if (bounds.length !=dimensions+1) {
					return false;
				}
				var expectedArraySize=1;
				for (var i=1; i < bounds.length; i++) {
					var dimension=bounds[i];
					if (isFinite(dimension)==false) {
						return false;
					}
					expectedArraySize=expectedArraySize * dimension;
				}
				expectedArraySize++;
				isSafeArray=(expectedArraySize==obj.length);
			}
			return isSafeArray;
		};
		return WebkitSafeArray;
	})();
	OSFWebkit.WebkitSafeArray=WebkitSafeArray;
})(OSFWebkit || (OSFWebkit={}));
var OSFWebkit;
(function (OSFWebkit) {
	var ScriptMessaging;
	(function (ScriptMessaging) {
		var scriptMessenger=null;
		function agaveHostCallback(callbackId, params) {
			scriptMessenger.agaveHostCallback(callbackId, params);
		}
		ScriptMessaging.agaveHostCallback=agaveHostCallback;
		function agaveHostEventCallback(callbackId, params) {
			scriptMessenger.agaveHostEventCallback(callbackId, params);
		}
		ScriptMessaging.agaveHostEventCallback=agaveHostEventCallback;
		function GetScriptMessenger() {
			if (scriptMessenger==null) {
				scriptMessenger=new WebkitScriptMessaging("OSF.ScriptMessaging.agaveHostCallback", "OSF.ScriptMessaging.agaveHostEventCallback");
			}
			return scriptMessenger;
		}
		ScriptMessaging.GetScriptMessenger=GetScriptMessenger;
		var EventHandlerCallback=(function () {
			function EventHandlerCallback(id, targetId, handler) {
				this.id=id;
				this.targetId=targetId;
				this.handler=handler;
			}
			return EventHandlerCallback;
		})();
		var WebkitScriptMessaging=(function () {
			function WebkitScriptMessaging(methodCallbackName, eventCallbackName) {
				this.callingIndex=0;
				this.callbackList={};
				this.eventHandlerList={};
				this.asyncMethodCallbackFunctionName=methodCallbackName;
				this.eventCallbackFunctionName=eventCallbackName;
				this.conversationId=WebkitScriptMessaging.getCurrentTimeMS().toString();
			}
			WebkitScriptMessaging.prototype.invokeMethod=function (handlerName, methodId, params, callback) {
				var messagingArgs={};
				this.postWebkitMessage(messagingArgs, handlerName, methodId, params, callback);
			};
			WebkitScriptMessaging.prototype.registerEvent=function (handlerName, methodId, dispId, targetId, handler, callback) {
				var messagingArgs={
					eventCallbackFunction: this.eventCallbackFunctionName
				};
				var hostArgs={
					id: dispId,
					targetId: targetId
				};
				var correlationId=this.postWebkitMessage(messagingArgs, handlerName, methodId, hostArgs, callback);
				this.eventHandlerList[correlationId]=new EventHandlerCallback(dispId, targetId, handler);
			};
			WebkitScriptMessaging.prototype.unregisterEvent=function (handlerName, methodId, dispId, targetId, callback) {
				var hostArgs={
					id: dispId,
					targetId: targetId
				};
				for (var key in this.eventHandlerList) {
					if (this.eventHandlerList.hasOwnProperty(key)) {
						var eventCallback=this.eventHandlerList[key];
						if (eventCallback.id==dispId && eventCallback.targetId==targetId) {
							delete this.eventHandlerList[key];
						}
					}
				}
				this.invokeMethod(handlerName, methodId, hostArgs, callback);
			};
			WebkitScriptMessaging.prototype.agaveHostCallback=function (callbackId, params) {
				var callbackFunction=this.callbackList[callbackId];
				if (callbackFunction) {
					var callbacksDone=callbackFunction(params);
					if (callbacksDone===undefined || callbacksDone===true) {
						delete this.callbackList[callbackId];
					}
				}
			};
			WebkitScriptMessaging.prototype.agaveHostEventCallback=function (callbackId, params) {
				var eventCallback=this.eventHandlerList[callbackId];
				if (eventCallback) {
					eventCallback.handler(params);
				}
			};
			WebkitScriptMessaging.prototype.postWebkitMessage=function (messagingArgs, handlerName, methodId, params, callback) {
				var correlationId=this.generateCorrelationId();
				this.callbackList[correlationId]=callback;
				messagingArgs.methodId=methodId;
				messagingArgs.params=params;
				messagingArgs.callbackId=correlationId;
				messagingArgs.callbackFunction=this.asyncMethodCallbackFunctionName;
				var invokePostMessage=function () {
					window.webkit.messageHandlers[handlerName].postMessage(JSON.stringify(messagingArgs));
				};
				var currentTimestamp=WebkitScriptMessaging.getCurrentTimeMS();
				if (this.lastMessageTimestamp==null || (currentTimestamp - this.lastMessageTimestamp >=WebkitScriptMessaging.MESSAGE_TIME_DELTA)) {
					invokePostMessage();
					this.lastMessageTimestamp=currentTimestamp;
				}
				else {
					this.lastMessageTimestamp+=WebkitScriptMessaging.MESSAGE_TIME_DELTA;
					setTimeout(function () {
						invokePostMessage();
					}, this.lastMessageTimestamp - currentTimestamp);
				}
				return correlationId;
			};
			WebkitScriptMessaging.prototype.generateCorrelationId=function () {
++this.callingIndex;
				return this.conversationId+this.callingIndex;
			};
			WebkitScriptMessaging.getCurrentTimeMS=function () {
				return (new Date).getTime();
			};
			WebkitScriptMessaging.MESSAGE_TIME_DELTA=10;
			return WebkitScriptMessaging;
		})();
		ScriptMessaging.WebkitScriptMessaging=WebkitScriptMessaging;
	})(ScriptMessaging=OSFWebkit.ScriptMessaging || (OSFWebkit.ScriptMessaging={}));
})(OSFWebkit || (OSFWebkit={}));
OSF.ScriptMessaging=OSFWebkit.ScriptMessaging;
var OSFWebkit;
(function (OSFWebkit) {
	OSFWebkit.MessageHandlerName="Agave";
	OSFWebkit.PopupMessageHandlerName="WefPopupHandler";
	(function (AppContextProperties) {
		AppContextProperties[AppContextProperties["Settings"]=0]="Settings";
		AppContextProperties[AppContextProperties["SolutionReferenceId"]=1]="SolutionReferenceId";
		AppContextProperties[AppContextProperties["AppType"]=2]="AppType";
		AppContextProperties[AppContextProperties["MajorVersion"]=3]="MajorVersion";
		AppContextProperties[AppContextProperties["MinorVersion"]=4]="MinorVersion";
		AppContextProperties[AppContextProperties["RevisionVersion"]=5]="RevisionVersion";
		AppContextProperties[AppContextProperties["APIVersionSequence"]=6]="APIVersionSequence";
		AppContextProperties[AppContextProperties["AppCapabilities"]=7]="AppCapabilities";
		AppContextProperties[AppContextProperties["APPUILocale"]=8]="APPUILocale";
		AppContextProperties[AppContextProperties["AppDataLocale"]=9]="AppDataLocale";
		AppContextProperties[AppContextProperties["BindingCount"]=10]="BindingCount";
		AppContextProperties[AppContextProperties["DocumentUrl"]=11]="DocumentUrl";
		AppContextProperties[AppContextProperties["ActivationMode"]=12]="ActivationMode";
		AppContextProperties[AppContextProperties["ControlIntegrationLevel"]=13]="ControlIntegrationLevel";
		AppContextProperties[AppContextProperties["SolutionToken"]=14]="SolutionToken";
		AppContextProperties[AppContextProperties["APISetVersion"]=15]="APISetVersion";
		AppContextProperties[AppContextProperties["CorrelationId"]=16]="CorrelationId";
		AppContextProperties[AppContextProperties["InstanceId"]=17]="InstanceId";
		AppContextProperties[AppContextProperties["TouchEnabled"]=18]="TouchEnabled";
		AppContextProperties[AppContextProperties["CommerceAllowed"]=19]="CommerceAllowed";
		AppContextProperties[AppContextProperties["RequirementMatrix"]=20]="RequirementMatrix";
		AppContextProperties[AppContextProperties["HostCustomMessage"]=21]="HostCustomMessage";
		AppContextProperties[AppContextProperties["HostFullVersion"]=22]="HostFullVersion";
	})(OSFWebkit.AppContextProperties || (OSFWebkit.AppContextProperties={}));
	var AppContextProperties=OSFWebkit.AppContextProperties;
	(function (MethodId) {
		MethodId[MethodId["Execute"]=1]="Execute";
		MethodId[MethodId["RegisterEvent"]=2]="RegisterEvent";
		MethodId[MethodId["UnregisterEvent"]=3]="UnregisterEvent";
		MethodId[MethodId["WriteSettings"]=4]="WriteSettings";
		MethodId[MethodId["GetContext"]=5]="GetContext";
		MethodId[MethodId["SendMessage"]=6]="SendMessage";
	})(OSFWebkit.MethodId || (OSFWebkit.MethodId={}));
	var MethodId=OSFWebkit.MethodId;
	var WebkitHostController=(function () {
		function WebkitHostController(hostScriptProxy) {
			this.hostScriptProxy=hostScriptProxy;
		}
		WebkitHostController.prototype.execute=function (id, params, callback) {
			var args=params;
			if (args==null) {
				args=[];
			}
			var hostParams={
				id: id,
				apiArgs: args
			};
			var agaveResponseCallback=function (payload) {
				if (callback) {
					var invokeArguments=[];
					if (OSF.OUtil.isArray(payload)) {
						for (var i=0; i < payload.length; i++) {
							var element=payload[i];
							if (OSF.OUtil.isArray(element)) {
								element=new OSFWebkit.WebkitSafeArray(element);
							}
							invokeArguments.unshift(element);
						}
					}
					return callback.apply(null, invokeArguments);
				}
			};
			this.hostScriptProxy.invokeMethod(OSF.Webkit.MessageHandlerName, OSF.Webkit.MethodId.Execute, hostParams, agaveResponseCallback);
		};
		WebkitHostController.prototype.registerEvent=function (id, targetId, handler, callback) {
			var agaveEventHandlerCallback=function (payload) {
				var safeArraySource=payload;
				var eventId=0;
				if (OSF.OUtil.isArray(payload) && payload.length >=2) {
					safeArraySource=payload[0];
					eventId=payload[1];
				}
				if (handler) {
					handler(eventId, new OSFWebkit.WebkitSafeArray(safeArraySource));
				}
			};
			var agaveResponseCallback=function (payload) {
				if (callback) {
					return callback(new OSFWebkit.WebkitSafeArray(payload));
				}
			};
			this.hostScriptProxy.registerEvent(OSF.Webkit.MessageHandlerName, OSF.Webkit.MethodId.RegisterEvent, id, targetId, agaveEventHandlerCallback, agaveResponseCallback);
		};
		WebkitHostController.prototype.unregisterEvent=function (id, targetId, callback) {
			var agaveResponseCallback=function (response) {
				return callback(new OSFWebkit.WebkitSafeArray(response));
			};
			this.hostScriptProxy.unregisterEvent(OSF.Webkit.MessageHandlerName, OSF.Webkit.MethodId.UnregisterEvent, id, targetId, agaveResponseCallback);
		};
		WebkitHostController.prototype.messageParent=function (params) {
			var message=params[Microsoft.Office.WebExtension.Parameters.MessageToParent];
			var messageObj={ dialogMessage: { messageType: OSF.DialogMessageType.DialogMessageReceived, messageContent: message } };
			window.opener.postMessage(JSON.stringify(messageObj), window.location.origin);
		};
		WebkitHostController.prototype.openDialog=function (id, targetId, handler, callback) {
			if (WebkitHostController.popup && !WebkitHostController.popup.closed) {
				callback(OSF.DDA.ErrorCodeManager.errorCodes.ooeDialogAlreadyOpened);
				return;
			}
			var magicWord="action=displayDialog";
			WebkitHostController.OpenDialogCallback=undefined;
			var fragmentSeparator='#';
			var callArgs=JSON.parse(targetId);
			var callUrl=callArgs.url;
			if (!callUrl) {
				return;
			}
			var urlParts=callUrl.split(fragmentSeparator);
			var seperator="?";
			if (urlParts[0].indexOf("?") > -1) {
				seperator="&";
			}
			var width=screen.width * callArgs.width / 100;
			var height=screen.height * callArgs.height / 100;
			var params="width="+width+", height="+height;
			urlParts[0]=urlParts[0].concat(seperator).concat(magicWord);
			var openUrl=urlParts.join(fragmentSeparator);
			WebkitHostController.popup=window.open(openUrl, "", params);
			function receiveMessage(event) {
				if (event.origin==window.location.origin) {
					try {
						var messageObj=JSON.parse(event.data);
						if (messageObj.dialogMessage) {
							handler(id, [OSF.DialogMessageType.DialogMessageReceived, messageObj.dialogMessage.messageContent]);
						}
					}
					catch (e) {
						OsfMsAjaxFactory.msAjaxDebug.trace("messages received cannot be handlered. Message:"+event.data);
					}
				}
			}
			function checkWindowClose() {
				try {
					if (WebkitHostController.popup==null || WebkitHostController.popup.closed) {
						window.clearInterval(WebkitHostController.interval);
						window.removeEventListener("message", receiveMessage);
						WebkitHostController.NotifyError=null;
						WebkitHostController.popup=null;
						handler(id, [OSF.DialogMessageType.DialogClosed]);
					}
				}
				catch (e) {
					OsfMsAjaxFactory.msAjaxDebug.trace("Error happened when popup window closed.");
				}
			}
			WebkitHostController.OpenDialogCallback=function (code) {
				if (code==OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
					window.addEventListener("message", receiveMessage);
					this.messageEventListener=receiveMessage;
					WebkitHostController.interval=window.setInterval(checkWindowClose, 1000);
					function notifyError(errorCode) {
						handler(id, [errorCode]);
					}
					WebkitHostController.NotifyError=notifyError;
				}
				callback(code);
			};
		};
		WebkitHostController.prototype.closeDialog=function (id, targetId, callback) {
			if (WebkitHostController.popup) {
				if (WebkitHostController.interval) {
					window.clearInterval(WebkitHostController.interval);
				}
				WebkitHostController.popup.close();
				WebkitHostController.popup=null;
				window.removeEventListener("message", this.messageEventListener);
				WebkitHostController.NotifyError=null;
				callback(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess);
			}
			else {
				callback(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);
			}
		};
		WebkitHostController.prototype.sendMessage=function (params) {
			var message=params[Microsoft.Office.WebExtension.Parameters.MessageContent];
			if (!isNaN(parseFloat(message)) && isFinite(message)) {
				message=message.toString();
			}
			this.hostScriptProxy.invokeMethod(OSF.Webkit.MessageHandlerName, OSF.Webkit.MethodId.SendMessage, message, null);
		};
		return WebkitHostController;
	})();
	OSFWebkit.WebkitHostController=WebkitHostController;
})(OSFWebkit || (OSFWebkit={}));
OSF.Webkit=OSFWebkit;
OSF.ClientHostController=new OSFWebkit.WebkitHostController(OSF.ScriptMessaging.GetScriptMessenger());
OSF.DDA.ClientSettingsManager={
	getSettingsExecuteMethod: function OSF_DDA_ClientSettingsManager$getSettingsExecuteMethod(hostDelegateMethod) {
		return function (args) {
			var status, response;
			var onComplete=function onComplete(status, response) {
				if (args.onReceiving) {
					args.onReceiving();
				}
				if (args.onComplete) {
					args.onComplete(status, response);
				}
			};
			try {
				hostDelegateMethod(args.hostCallArgs, args.onCalling, onComplete);
			}
			catch (ex) {
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
				response={ name: Strings.OfficeOM.L_InternalError, message: ex };
				onComplete(status, response);
			}
		};
	},
	read: function OSF_DDA_ClientSettingsManager$read(onCalling, onComplete) {
		var keys=[];
		var values=[];
		if (onCalling) {
			onCalling();
		}
		var initializationHelper=OSF._OfficeAppFactory.getInitializationHelper();
		var onReceivedContext=function onReceivedContext(appContext) {
			if (onComplete) {
				onComplete(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess, appContext.get_settings());
			}
		};
		initializationHelper.getAppContext(null, onReceivedContext);
	},
	write: function OSF_DDA_ClientSettingsManager$write(serializedSettings, overwriteIfStale, onCalling, onComplete) {
		var hostParams={};
		var keys=[];
		var values=[];
		for (var key in serializedSettings) {
			keys.push(key);
			values.push(serializedSettings[key]);
		}
		hostParams["keys"]=keys;
		hostParams["values"]=values;
		if (onCalling) {
			onCalling();
		}
		var onWriteCompleted=function onWriteCompleted(status) {
			if (onComplete) {
				onComplete(status[0], null);
			}
		};
		OSF.ScriptMessaging.GetScriptMessenger().invokeMethod(OSF.Webkit.MessageHandlerName, OSF.Webkit.MethodId.WriteSettings, hostParams, onWriteCompleted);
	}
};
OSF.InitializationHelper.prototype.initializeSettings=function OSF_InitializationHelper$initializeSettings(appContext, refreshSupported) {
	var serializedSettings=appContext.get_settings();
	var settings=this.deserializeSettings(serializedSettings, refreshSupported);
	return settings;
};
OSF.InitializationHelper.prototype.getAppContext=function OSF_InitializationHelper$getAppContext(wnd, gotAppContext) {
	var getInvocationCallback=function OSF_InitializationHelper_getAppContextAsync$getInvocationCallbackWebApp(appContext) {
		var returnedContext;
		var appContextProperties=OSF.Webkit.AppContextProperties;
		var appType=appContext[appContextProperties.AppType];
		var hostSettings=appContext[appContextProperties.Settings];
		var serializedSettings={};
		var keys=hostSettings[0];
		var values=hostSettings[1];
		for (var index=0; index < keys.length; index++) {
			serializedSettings[keys[index]]=values[index];
		}
		var id=appContext[appContextProperties.SolutionReferenceId];
		var version=appContext[appContextProperties.MajorVersion];
		var minorVersion=appContext[appContextProperties.MinorVersion];
		var clientMode=appContext[appContextProperties.AppCapabilities];
		var UILocale=appContext[appContextProperties.APPUILocale];
		var dataLocale=appContext[appContextProperties.AppDataLocale];
		var docUrl=appContext[appContextProperties.DocumentUrl];
		var reason=appContext[appContextProperties.ActivationMode];
		var osfControlType=appContext[appContextProperties.ControlIntegrationLevel];
		var eToken=appContext[appContextProperties.SolutionToken];
		eToken=eToken ? eToken.toString() : "";
		var correlationId=appContext[appContextProperties.CorrelationId];
		var appInstanceId=appContext[appContextProperties.InstanceId];
		var touchEnabled=appContext[appContextProperties.TouchEnabled];
		var commerceAllowed=appContext[appContextProperties.CommerceAllowed];
		var requirementMatrix=appContext[appContextProperties.RequirementMatrix];
		var hostCustomMessage=appContext[appContextProperties.HostCustomMessage];
		var hostFullVersion=appContext[appContextProperties.HostFullVersion];
		returnedContext=new OSF.OfficeAppContext(id, appType, version, UILocale, dataLocale, docUrl, clientMode, serializedSettings, reason, osfControlType, eToken, correlationId, appInstanceId, touchEnabled, commerceAllowed, minorVersion, requirementMatrix, hostCustomMessage, hostFullVersion);
		if (OSF.AppTelemetry) {
			OSF.AppTelemetry.initialize(returnedContext);
		}
		gotAppContext(returnedContext);
	};
	var handler;
	if (this._hostInfo.isDialog) {
		handler=OSF.Webkit.PopupMessageHandlerName;
	}
	else {
		handler=OSF.Webkit.MessageHandlerName;
	}
	OSF.ScriptMessaging.GetScriptMessenger().invokeMethod(handler, OSF.Webkit.MethodId.GetContext, [], getInvocationCallback);
};
var OSFLog;
(function (OSFLog) {
	var BaseUsageData=(function () {
		function BaseUsageData(table) {
			this._table=table;
			this._fields={};
		}
		Object.defineProperty(BaseUsageData.prototype, "Fields", {
			get: function () {
				return this._fields;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(BaseUsageData.prototype, "Table", {
			get: function () {
				return this._table;
			},
			enumerable: true,
			configurable: true
		});
		BaseUsageData.prototype.SerializeFields=function () {
		};
		BaseUsageData.prototype.SetSerializedField=function (key, value) {
			if (typeof (value) !=="undefined" && value !==null) {
				this._serializedFields[key]=value.toString();
			}
		};
		BaseUsageData.prototype.SerializeRow=function () {
			this._serializedFields={};
			this.SetSerializedField("Table", this._table);
			this.SerializeFields();
			return JSON.stringify(this._serializedFields);
		};
		return BaseUsageData;
	})();
	OSFLog.BaseUsageData=BaseUsageData;
	var AppActivatedUsageData=(function (_super) {
		__extends(AppActivatedUsageData, _super);
		function AppActivatedUsageData() {
			_super.call(this, "AppActivated");
		}
		Object.defineProperty(AppActivatedUsageData.prototype, "CorrelationId", {
			get: function () { return this.Fields["CorrelationId"]; },
			set: function (value) { this.Fields["CorrelationId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "SessionId", {
			get: function () { return this.Fields["SessionId"]; },
			set: function (value) { this.Fields["SessionId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "AppId", {
			get: function () { return this.Fields["AppId"]; },
			set: function (value) { this.Fields["AppId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "AppInstanceId", {
			get: function () { return this.Fields["AppInstanceId"]; },
			set: function (value) { this.Fields["AppInstanceId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "AppURL", {
			get: function () { return this.Fields["AppURL"]; },
			set: function (value) { this.Fields["AppURL"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "AssetId", {
			get: function () { return this.Fields["AssetId"]; },
			set: function (value) { this.Fields["AssetId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "Browser", {
			get: function () { return this.Fields["Browser"]; },
			set: function (value) { this.Fields["Browser"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "UserId", {
			get: function () { return this.Fields["UserId"]; },
			set: function (value) { this.Fields["UserId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "Host", {
			get: function () { return this.Fields["Host"]; },
			set: function (value) { this.Fields["Host"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "HostVersion", {
			get: function () { return this.Fields["HostVersion"]; },
			set: function (value) { this.Fields["HostVersion"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "ClientId", {
			get: function () { return this.Fields["ClientId"]; },
			set: function (value) { this.Fields["ClientId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "AppSizeWidth", {
			get: function () { return this.Fields["AppSizeWidth"]; },
			set: function (value) { this.Fields["AppSizeWidth"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "AppSizeHeight", {
			get: function () { return this.Fields["AppSizeHeight"]; },
			set: function (value) { this.Fields["AppSizeHeight"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "Message", {
			get: function () { return this.Fields["Message"]; },
			set: function (value) { this.Fields["Message"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "DocUrl", {
			get: function () { return this.Fields["DocUrl"]; },
			set: function (value) { this.Fields["DocUrl"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "OfficeJSVersion", {
			get: function () { return this.Fields["OfficeJSVersion"]; },
			set: function (value) { this.Fields["OfficeJSVersion"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "HostJSVersion", {
			get: function () { return this.Fields["HostJSVersion"]; },
			set: function (value) { this.Fields["HostJSVersion"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "WacHostEnvironment", {
			get: function () { return this.Fields["WacHostEnvironment"]; },
			set: function (value) { this.Fields["WacHostEnvironment"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppActivatedUsageData.prototype, "IsFromWacAutomation", {
			get: function () { return this.Fields["IsFromWacAutomation"]; },
			set: function (value) { this.Fields["IsFromWacAutomation"]=value; },
			enumerable: true,
			configurable: true
		});
		AppActivatedUsageData.prototype.SerializeFields=function () {
			this.SetSerializedField("CorrelationId", this.CorrelationId);
			this.SetSerializedField("SessionId", this.SessionId);
			this.SetSerializedField("AppId", this.AppId);
			this.SetSerializedField("AppInstanceId", this.AppInstanceId);
			this.SetSerializedField("AppURL", this.AppURL);
			this.SetSerializedField("AssetId", this.AssetId);
			this.SetSerializedField("Browser", this.Browser);
			this.SetSerializedField("UserId", this.UserId);
			this.SetSerializedField("Host", this.Host);
			this.SetSerializedField("HostVersion", this.HostVersion);
			this.SetSerializedField("ClientId", this.ClientId);
			this.SetSerializedField("AppSizeWidth", this.AppSizeWidth);
			this.SetSerializedField("AppSizeHeight", this.AppSizeHeight);
			this.SetSerializedField("Message", this.Message);
			this.SetSerializedField("DocUrl", this.DocUrl);
			this.SetSerializedField("OfficeJSVersion", this.OfficeJSVersion);
			this.SetSerializedField("HostJSVersion", this.HostJSVersion);
			this.SetSerializedField("WacHostEnvironment", this.WacHostEnvironment);
			this.SetSerializedField("IsFromWacAutomation", this.IsFromWacAutomation);
		};
		return AppActivatedUsageData;
	})(BaseUsageData);
	OSFLog.AppActivatedUsageData=AppActivatedUsageData;
	var ScriptLoadUsageData=(function (_super) {
		__extends(ScriptLoadUsageData, _super);
		function ScriptLoadUsageData() {
			_super.call(this, "ScriptLoad");
		}
		Object.defineProperty(ScriptLoadUsageData.prototype, "CorrelationId", {
			get: function () { return this.Fields["CorrelationId"]; },
			set: function (value) { this.Fields["CorrelationId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ScriptLoadUsageData.prototype, "SessionId", {
			get: function () { return this.Fields["SessionId"]; },
			set: function (value) { this.Fields["SessionId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ScriptLoadUsageData.prototype, "ScriptId", {
			get: function () { return this.Fields["ScriptId"]; },
			set: function (value) { this.Fields["ScriptId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ScriptLoadUsageData.prototype, "StartTime", {
			get: function () { return this.Fields["StartTime"]; },
			set: function (value) { this.Fields["StartTime"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ScriptLoadUsageData.prototype, "ResponseTime", {
			get: function () { return this.Fields["ResponseTime"]; },
			set: function (value) { this.Fields["ResponseTime"]=value; },
			enumerable: true,
			configurable: true
		});
		ScriptLoadUsageData.prototype.SerializeFields=function () {
			this.SetSerializedField("CorrelationId", this.CorrelationId);
			this.SetSerializedField("SessionId", this.SessionId);
			this.SetSerializedField("ScriptId", this.ScriptId);
			this.SetSerializedField("StartTime", this.StartTime);
			this.SetSerializedField("ResponseTime", this.ResponseTime);
		};
		return ScriptLoadUsageData;
	})(BaseUsageData);
	OSFLog.ScriptLoadUsageData=ScriptLoadUsageData;
	var AppClosedUsageData=(function (_super) {
		__extends(AppClosedUsageData, _super);
		function AppClosedUsageData() {
			_super.call(this, "AppClosed");
		}
		Object.defineProperty(AppClosedUsageData.prototype, "CorrelationId", {
			get: function () { return this.Fields["CorrelationId"]; },
			set: function (value) { this.Fields["CorrelationId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppClosedUsageData.prototype, "SessionId", {
			get: function () { return this.Fields["SessionId"]; },
			set: function (value) { this.Fields["SessionId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppClosedUsageData.prototype, "FocusTime", {
			get: function () { return this.Fields["FocusTime"]; },
			set: function (value) { this.Fields["FocusTime"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppClosedUsageData.prototype, "AppSizeFinalWidth", {
			get: function () { return this.Fields["AppSizeFinalWidth"]; },
			set: function (value) { this.Fields["AppSizeFinalWidth"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppClosedUsageData.prototype, "AppSizeFinalHeight", {
			get: function () { return this.Fields["AppSizeFinalHeight"]; },
			set: function (value) { this.Fields["AppSizeFinalHeight"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppClosedUsageData.prototype, "OpenTime", {
			get: function () { return this.Fields["OpenTime"]; },
			set: function (value) { this.Fields["OpenTime"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppClosedUsageData.prototype, "CloseMethod", {
			get: function () { return this.Fields["CloseMethod"]; },
			set: function (value) { this.Fields["CloseMethod"]=value; },
			enumerable: true,
			configurable: true
		});
		AppClosedUsageData.prototype.SerializeFields=function () {
			this.SetSerializedField("CorrelationId", this.CorrelationId);
			this.SetSerializedField("SessionId", this.SessionId);
			this.SetSerializedField("FocusTime", this.FocusTime);
			this.SetSerializedField("AppSizeFinalWidth", this.AppSizeFinalWidth);
			this.SetSerializedField("AppSizeFinalHeight", this.AppSizeFinalHeight);
			this.SetSerializedField("OpenTime", this.OpenTime);
			this.SetSerializedField("CloseMethod", this.CloseMethod);
		};
		return AppClosedUsageData;
	})(BaseUsageData);
	OSFLog.AppClosedUsageData=AppClosedUsageData;
	var APIUsageUsageData=(function (_super) {
		__extends(APIUsageUsageData, _super);
		function APIUsageUsageData() {
			_super.call(this, "APIUsage");
		}
		Object.defineProperty(APIUsageUsageData.prototype, "CorrelationId", {
			get: function () { return this.Fields["CorrelationId"]; },
			set: function (value) { this.Fields["CorrelationId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(APIUsageUsageData.prototype, "SessionId", {
			get: function () { return this.Fields["SessionId"]; },
			set: function (value) { this.Fields["SessionId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(APIUsageUsageData.prototype, "APIType", {
			get: function () { return this.Fields["APIType"]; },
			set: function (value) { this.Fields["APIType"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(APIUsageUsageData.prototype, "APIID", {
			get: function () { return this.Fields["APIID"]; },
			set: function (value) { this.Fields["APIID"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(APIUsageUsageData.prototype, "Parameters", {
			get: function () { return this.Fields["Parameters"]; },
			set: function (value) { this.Fields["Parameters"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(APIUsageUsageData.prototype, "ResponseTime", {
			get: function () { return this.Fields["ResponseTime"]; },
			set: function (value) { this.Fields["ResponseTime"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(APIUsageUsageData.prototype, "ErrorType", {
			get: function () { return this.Fields["ErrorType"]; },
			set: function (value) { this.Fields["ErrorType"]=value; },
			enumerable: true,
			configurable: true
		});
		APIUsageUsageData.prototype.SerializeFields=function () {
			this.SetSerializedField("CorrelationId", this.CorrelationId);
			this.SetSerializedField("SessionId", this.SessionId);
			this.SetSerializedField("APIType", this.APIType);
			this.SetSerializedField("APIID", this.APIID);
			this.SetSerializedField("Parameters", this.Parameters);
			this.SetSerializedField("ResponseTime", this.ResponseTime);
			this.SetSerializedField("ErrorType", this.ErrorType);
		};
		return APIUsageUsageData;
	})(BaseUsageData);
	OSFLog.APIUsageUsageData=APIUsageUsageData;
	var AppInitializationUsageData=(function (_super) {
		__extends(AppInitializationUsageData, _super);
		function AppInitializationUsageData() {
			_super.call(this, "AppInitialization");
		}
		Object.defineProperty(AppInitializationUsageData.prototype, "CorrelationId", {
			get: function () { return this.Fields["CorrelationId"]; },
			set: function (value) { this.Fields["CorrelationId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppInitializationUsageData.prototype, "SessionId", {
			get: function () { return this.Fields["SessionId"]; },
			set: function (value) { this.Fields["SessionId"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppInitializationUsageData.prototype, "SuccessCode", {
			get: function () { return this.Fields["SuccessCode"]; },
			set: function (value) { this.Fields["SuccessCode"]=value; },
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(AppInitializationUsageData.prototype, "Message", {
			get: function () { return this.Fields["Message"]; },
			set: function (value) { this.Fields["Message"]=value; },
			enumerable: true,
			configurable: true
		});
		AppInitializationUsageData.prototype.SerializeFields=function () {
			this.SetSerializedField("CorrelationId", this.CorrelationId);
			this.SetSerializedField("SessionId", this.SessionId);
			this.SetSerializedField("SuccessCode", this.SuccessCode);
			this.SetSerializedField("Message", this.Message);
		};
		return AppInitializationUsageData;
	})(BaseUsageData);
	OSFLog.AppInitializationUsageData=AppInitializationUsageData;
})(OSFLog || (OSFLog={}));
var Logger;
(function (Logger) {
	"use strict";
	(function (TraceLevel) {
		TraceLevel[TraceLevel["info"]=0]="info";
		TraceLevel[TraceLevel["warning"]=1]="warning";
		TraceLevel[TraceLevel["error"]=2]="error";
	})(Logger.TraceLevel || (Logger.TraceLevel={}));
	var TraceLevel=Logger.TraceLevel;
	(function (SendFlag) {
		SendFlag[SendFlag["none"]=0]="none";
		SendFlag[SendFlag["flush"]=1]="flush";
	})(Logger.SendFlag || (Logger.SendFlag={}));
	var SendFlag=Logger.SendFlag;
	function allowUploadingData() {
	}
	Logger.allowUploadingData=allowUploadingData;
	function sendLog(traceLevel, message, flag) {
	}
	Logger.sendLog=sendLog;
	function creatULSEndpoint() {
		try {
			return new ULSEndpointProxy();
		}
		catch (e) {
			return null;
		}
	}
	var ULSEndpointProxy=(function () {
		function ULSEndpointProxy() {
		}
		ULSEndpointProxy.prototype.writeLog=function (log) {
		};
		ULSEndpointProxy.prototype.loadProxyFrame=function () {
		};
		return ULSEndpointProxy;
	})();
	if (!OSF.Logger) {
		OSF.Logger=Logger;
	}
	Logger.ulsEndpoint=creatULSEndpoint();
})(Logger || (Logger={}));
var OSFAriaLogger;
(function (OSFAriaLogger) {
	var AriaLogger=(function () {
		function AriaLogger() {
		}
		AriaLogger.prototype.getAriaCDNLocation=function () {
			return (OSF._OfficeAppFactory.getLoadScriptHelper().getOfficeJsBasePath()+"/ariatelemetry/aria-web-telemetry.js");
		};
		AriaLogger.getInstance=function () {
			if (AriaLogger.AriaLoggerObj===undefined) {
				AriaLogger.AriaLoggerObj=new AriaLogger();
			}
			return AriaLogger.AriaLoggerObj;
		};
		AriaLogger.prototype.isIUsageData=function (arg) {
			return arg["Fields"] !==undefined;
		};
		AriaLogger.prototype.loadAriaScriptAndLog=function (tableName, telemetryData) {
			var startAfterMs=1000;
			OSF.OUtil.loadScript(this.getAriaCDNLocation(), function () {
				try {
					if (!this.ALogger) {
						var OfficeExtensibilityTenantID="db334b301e7b474db5e0f02f07c51a47-a1b5bc36-1bbe-482f-a64a-c2d9cb606706-7439";
						this.ALogger=AWTLogManager.initialize(OfficeExtensibilityTenantID);
					}
					var eventProperties=new AWTEventProperties();
					eventProperties.setName("Office.Extensibility.OfficeJS."+tableName);
					for (var key in telemetryData) {
						if (key.toLowerCase() !=="table") {
							eventProperties.setProperty(key, telemetryData[key]);
						}
					}
					var today=new Date();
					eventProperties.setProperty("Date", today.toISOString());
					this.ALogger.logEvent(eventProperties);
				}
				catch (e) {
				}
			}, startAfterMs);
		};
		AriaLogger.prototype.logData=function (data) {
			if (this.isIUsageData(data)) {
				this.loadAriaScriptAndLog(data["Table"], data["Fields"]);
			}
			else {
				this.loadAriaScriptAndLog(data["Table"], data);
			}
		};
		return AriaLogger;
	})();
	OSFAriaLogger.AriaLogger=AriaLogger;
})(OSFAriaLogger || (OSFAriaLogger={}));
var OSFAppTelemetry;
(function (OSFAppTelemetry) {
	"use strict";
	var appInfo;
	var sessionId=OSF.OUtil.Guid.generateNewGuid();
	var osfControlAppCorrelationId="";
	var omexDomainRegex=new RegExp("^https?://store\\.office(ppe|-int)?\\.com/", "i");
	OSFAppTelemetry.enableTelemetry=true;
	;
	var AppInfo=(function () {
		function AppInfo() {
		}
		return AppInfo;
	})();
	var Event=(function () {
		function Event(name, handler) {
			this.name=name;
			this.handler=handler;
		}
		return Event;
	})();
	var AppStorage=(function () {
		function AppStorage() {
			this.clientIDKey="Office API client";
			this.logIdSetKey="Office App Log Id Set";
		}
		AppStorage.prototype.getClientId=function () {
			var clientId=this.getValue(this.clientIDKey);
			if (!clientId || clientId.length <=0 || clientId.length > 40) {
				clientId=OSF.OUtil.Guid.generateNewGuid();
				this.setValue(this.clientIDKey, clientId);
			}
			return clientId;
		};
		AppStorage.prototype.saveLog=function (logId, log) {
			var logIdSet=this.getValue(this.logIdSetKey);
			logIdSet=((logIdSet && logIdSet.length > 0) ? (logIdSet+";") : "")+logId;
			this.setValue(this.logIdSetKey, logIdSet);
			this.setValue(logId, log);
		};
		AppStorage.prototype.enumerateLog=function (callback, clean) {
			var logIdSet=this.getValue(this.logIdSetKey);
			if (logIdSet) {
				var ids=logIdSet.split(";");
				for (var id in ids) {
					var logId=ids[id];
					var log=this.getValue(logId);
					if (log) {
						if (callback) {
							callback(logId, log);
						}
						if (clean) {
							this.remove(logId);
						}
					}
				}
				if (clean) {
					this.remove(this.logIdSetKey);
				}
			}
		};
		AppStorage.prototype.getValue=function (key) {
			var osfLocalStorage=OSF.OUtil.getLocalStorage();
			var value="";
			if (osfLocalStorage) {
				value=osfLocalStorage.getItem(key);
			}
			return value;
		};
		AppStorage.prototype.setValue=function (key, value) {
			var osfLocalStorage=OSF.OUtil.getLocalStorage();
			if (osfLocalStorage) {
				osfLocalStorage.setItem(key, value);
			}
		};
		AppStorage.prototype.remove=function (key) {
			var osfLocalStorage=OSF.OUtil.getLocalStorage();
			if (osfLocalStorage) {
				try {
					osfLocalStorage.removeItem(key);
				}
				catch (ex) {
				}
			}
		};
		return AppStorage;
	})();
	var AppLogger=(function () {
		function AppLogger() {
		}
		AppLogger.prototype.LogData=function (data) {
			if (!OSFAppTelemetry.enableTelemetry) {
				return;
			}
			try {
				OSFAriaLogger.AriaLogger.getInstance().logData(data);
			}
			catch (e) {
			}
		};
		AppLogger.prototype.LogRawData=function (log) {
			if (!OSFAppTelemetry.enableTelemetry) {
				return;
			}
			try {
				OSFAriaLogger.AriaLogger.getInstance().logData(JSON.parse(log));
			}
			catch (e) {
			}
		};
		return AppLogger;
	})();
	function trimStringToLowerCase(input) {
		if (input) {
			input=input.replace(/[{}]/g, "").toLowerCase();
		}
		return (input || "");
	}
	function initialize(context) {
		if (!OSFAppTelemetry.enableTelemetry) {
			return;
		}
		if (appInfo) {
			return;
		}
		appInfo=new AppInfo();
		if (context.get_hostFullVersion()) {
			appInfo.hostVersion=context.get_hostFullVersion();
		}
		else {
			appInfo.hostVersion=context.get_appVersion();
		}
		appInfo.appId=context.get_id();
		appInfo.host=context.get_appName();
		appInfo.browser=window.navigator.userAgent;
		appInfo.correlationId=trimStringToLowerCase(context.get_correlationId());
		appInfo.clientId=(new AppStorage()).getClientId();
		appInfo.appInstanceId=context.get_appInstanceId();
		if (appInfo.appInstanceId) {
			appInfo.appInstanceId=appInfo.appInstanceId.replace(/[{}]/g, "").toLowerCase();
		}
		appInfo.message=context.get_hostCustomMessage();
		appInfo.officeJSVersion=OSF.ConstantNames.FileVersion;
		appInfo.hostJSVersion="16.0.9124.1000";
		if (context._wacHostEnvironment) {
			appInfo.wacHostEnvironment=context._wacHostEnvironment;
		}
		if (context._isFromWacAutomation !==undefined && context._isFromWacAutomation !==null) {
			appInfo.isFromWacAutomation=context._isFromWacAutomation.toString().toLowerCase();
		}
		var docUrl=context.get_docUrl();
		appInfo.docUrl=omexDomainRegex.test(docUrl) ? docUrl : "";
		var url=location.href;
		if (url) {
			url=url.split("?")[0].split("#")[0];
		}
		appInfo.appURL=url;
		(function getUserIdAndAssetIdFromToken(token, appInfo) {
			var xmlContent;
			var parser;
			var xmlDoc;
			appInfo.assetId="";
			appInfo.userId="";
			try {
				xmlContent=decodeURIComponent(token);
				parser=new DOMParser();
				xmlDoc=parser.parseFromString(xmlContent, "text/xml");
				var cidNode=xmlDoc.getElementsByTagName("t")[0].attributes.getNamedItem("cid");
				var oidNode=xmlDoc.getElementsByTagName("t")[0].attributes.getNamedItem("oid");
				if (cidNode && cidNode.nodeValue) {
					appInfo.userId=cidNode.nodeValue;
				}
				else if (oidNode && oidNode.nodeValue) {
					appInfo.userId=oidNode.nodeValue;
				}
				appInfo.assetId=xmlDoc.getElementsByTagName("t")[0].attributes.getNamedItem("aid").nodeValue;
			}
			catch (e) {
			}
			finally {
				xmlContent=null;
				xmlDoc=null;
				parser=null;
			}
		})(context.get_eToken(), appInfo);
		(function handleLifecycle() {
			var startTime=new Date();
			var lastFocus=null;
			var focusTime=0;
			var finished=false;
			var adjustFocusTime=function () {
				if (document.hasFocus()) {
					if (lastFocus==null) {
						lastFocus=new Date();
					}
				}
				else if (lastFocus) {
					focusTime+=Math.abs((new Date()).getTime() - lastFocus.getTime());
					lastFocus=null;
				}
			};
			var eventList=[];
			eventList.push(new Event("focus", adjustFocusTime));
			eventList.push(new Event("blur", adjustFocusTime));
			eventList.push(new Event("focusout", adjustFocusTime));
			eventList.push(new Event("focusin", adjustFocusTime));
			var exitFunction=function () {
				for (var i=0; i < eventList.length; i++) {
					OSF.OUtil.removeEventListener(window, eventList[i].name, eventList[i].handler);
				}
				eventList.length=0;
				if (!finished) {
					if (document.hasFocus() && lastFocus) {
						focusTime+=Math.abs((new Date()).getTime() - lastFocus.getTime());
						lastFocus=null;
					}
					OSFAppTelemetry.onAppClosed(Math.abs((new Date()).getTime() - startTime.getTime()), focusTime);
					finished=true;
				}
			};
			eventList.push(new Event("beforeunload", exitFunction));
			eventList.push(new Event("unload", exitFunction));
			for (var i=0; i < eventList.length; i++) {
				OSF.OUtil.addEventListener(window, eventList[i].name, eventList[i].handler);
			}
			adjustFocusTime();
		})();
		OSFAppTelemetry.onAppActivated();
	}
	OSFAppTelemetry.initialize=initialize;
	function onAppActivated() {
		if (!appInfo) {
			return;
		}
		(new AppStorage()).enumerateLog(function (id, log) { return (new AppLogger()).LogRawData(log); }, true);
		var data=new OSFLog.AppActivatedUsageData();
		data.SessionId=sessionId;
		data.AppId=appInfo.appId;
		data.AssetId=appInfo.assetId;
		data.AppURL=appInfo.appURL;
		data.UserId=appInfo.userId;
		data.ClientId=appInfo.clientId;
		data.Browser=appInfo.browser;
		data.Host=appInfo.host;
		data.HostVersion=appInfo.hostVersion;
		data.CorrelationId=trimStringToLowerCase(appInfo.correlationId);
		data.AppSizeWidth=window.innerWidth;
		data.AppSizeHeight=window.innerHeight;
		data.AppInstanceId=appInfo.appInstanceId;
		data.Message=appInfo.message;
		data.DocUrl=appInfo.docUrl;
		data.OfficeJSVersion=appInfo.officeJSVersion;
		data.HostJSVersion=appInfo.hostJSVersion;
		if (appInfo.wacHostEnvironment) {
			data.WacHostEnvironment=appInfo.wacHostEnvironment;
		}
		if (appInfo.isFromWacAutomation !==undefined && appInfo.isFromWacAutomation !==null) {
			data.IsFromWacAutomation=appInfo.isFromWacAutomation;
		}
		(new AppLogger()).LogData(data);
	}
	OSFAppTelemetry.onAppActivated=onAppActivated;
	function onScriptDone(scriptId, msStartTime, msResponseTime, appCorrelationId) {
		var data=new OSFLog.ScriptLoadUsageData();
		data.CorrelationId=trimStringToLowerCase(appCorrelationId);
		data.SessionId=sessionId;
		data.ScriptId=scriptId;
		data.StartTime=msStartTime;
		data.ResponseTime=msResponseTime;
		(new AppLogger()).LogData(data);
	}
	OSFAppTelemetry.onScriptDone=onScriptDone;
	function onCallDone(apiType, id, parameters, msResponseTime, errorType) {
		if (!appInfo) {
			return;
		}
		var data=new OSFLog.APIUsageUsageData();
		data.CorrelationId=trimStringToLowerCase(osfControlAppCorrelationId);
		data.SessionId=sessionId;
		data.APIType=apiType;
		data.APIID=id;
		data.Parameters=parameters;
		data.ResponseTime=msResponseTime;
		data.ErrorType=errorType;
		(new AppLogger()).LogData(data);
	}
	OSFAppTelemetry.onCallDone=onCallDone;
	;
	function onMethodDone(id, args, msResponseTime, errorType) {
		var parameters=null;
		if (args) {
			if (typeof args=="number") {
				parameters=String(args);
			}
			else if (typeof args==="object") {
				for (var index in args) {
					if (parameters !==null) {
						parameters+=",";
					}
					else {
						parameters="";
					}
					if (typeof args[index]=="number") {
						parameters+=String(args[index]);
					}
				}
			}
			else {
				parameters="";
			}
		}
		OSF.AppTelemetry.onCallDone("method", id, parameters, msResponseTime, errorType);
	}
	OSFAppTelemetry.onMethodDone=onMethodDone;
	function onPropertyDone(propertyName, msResponseTime) {
		OSF.AppTelemetry.onCallDone("property", -1, propertyName, msResponseTime);
	}
	OSFAppTelemetry.onPropertyDone=onPropertyDone;
	function onEventDone(id, errorType) {
		OSF.AppTelemetry.onCallDone("event", id, null, 0, errorType);
	}
	OSFAppTelemetry.onEventDone=onEventDone;
	function onRegisterDone(register, id, msResponseTime, errorType) {
		OSF.AppTelemetry.onCallDone(register ? "registerevent" : "unregisterevent", id, null, msResponseTime, errorType);
	}
	OSFAppTelemetry.onRegisterDone=onRegisterDone;
	function onAppClosed(openTime, focusTime) {
		if (!appInfo) {
			return;
		}
		var data=new OSFLog.AppClosedUsageData();
		data.CorrelationId=trimStringToLowerCase(osfControlAppCorrelationId);
		data.SessionId=sessionId;
		data.FocusTime=focusTime;
		data.OpenTime=openTime;
		data.AppSizeFinalWidth=window.innerWidth;
		data.AppSizeFinalHeight=window.innerHeight;
		(new AppStorage()).saveLog(sessionId, data.SerializeRow());
	}
	OSFAppTelemetry.onAppClosed=onAppClosed;
	function setOsfControlAppCorrelationId(correlationId) {
		osfControlAppCorrelationId=trimStringToLowerCase(correlationId);
	}
	OSFAppTelemetry.setOsfControlAppCorrelationId=setOsfControlAppCorrelationId;
	function doAppInitializationLogging(isException, message) {
		var data=new OSFLog.AppInitializationUsageData();
		data.CorrelationId=trimStringToLowerCase(osfControlAppCorrelationId);
		data.SessionId=sessionId;
		data.SuccessCode=isException ? 1 : 0;
		data.Message=message;
		(new AppLogger()).LogData(data);
	}
	OSFAppTelemetry.doAppInitializationLogging=doAppInitializationLogging;
	function logAppCommonMessage(message) {
		doAppInitializationLogging(false, message);
	}
	OSFAppTelemetry.logAppCommonMessage=logAppCommonMessage;
	function logAppException(errorMessage) {
		doAppInitializationLogging(true, errorMessage);
	}
	OSFAppTelemetry.logAppException=logAppException;
	OSF.AppTelemetry=OSFAppTelemetry;
})(OSFAppTelemetry || (OSFAppTelemetry={}));
Microsoft.Office.WebExtension.FileType={
	Text: "text",
	Compressed: "compressed",
	Pdf: "pdf"
};
OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors, {
	FileProperties: "FileProperties",
	FileSliceProperties: "FileSliceProperties"
});
OSF.DDA.FileProperties={
	Handle: "FileHandle",
	FileSize: "FileSize",
	SliceSize: Microsoft.Office.WebExtension.Parameters.SliceSize
};
OSF.DDA.File=function OSF_DDA_File(handle, fileSize, sliceSize) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"size": {
			value: fileSize
		},
		"sliceCount": {
			value: Math.ceil(fileSize / sliceSize)
		}
	});
	var privateState={};
	privateState[OSF.DDA.FileProperties.Handle]=handle;
	privateState[OSF.DDA.FileProperties.SliceSize]=sliceSize;
	var am=OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(this, [
		am.GetDocumentCopyChunkAsync,
		am.ReleaseDocumentCopyAsync
	], privateState);
};
OSF.DDA.FileSliceOffset="fileSliceoffset";
OSF.DDA.AsyncMethodNames.addNames({
	GetDocumentCopyAsync: "getFileAsync",
	GetDocumentCopyChunkAsync: "getSliceAsync",
	ReleaseDocumentCopyAsync: "closeAsync"
});
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.GetDocumentCopyAsync,
	requiredArguments: [
		{
			"name": Microsoft.Office.WebExtension.Parameters.FileType,
			"enum": Microsoft.Office.WebExtension.FileType
		}
	],
	supportedOptions: [
		{
			name: Microsoft.Office.WebExtension.Parameters.SliceSize,
			value: {
				"types": ["number"],
				"defaultValue": 64 * 1024
			}
		}
	],
	checkCallArgs: function (callArgs, caller, stateInfo) {
		var sliceSize=callArgs[Microsoft.Office.WebExtension.Parameters.SliceSize];
		if (sliceSize <=0 || sliceSize > (4 * 1024 * 1024)) {
			throw OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSliceSize;
		}
		return callArgs;
	},
	onSucceeded: function (fileDescriptor, caller, callArgs) {
		return new OSF.DDA.File(fileDescriptor[OSF.DDA.FileProperties.Handle], fileDescriptor[OSF.DDA.FileProperties.FileSize], callArgs[Microsoft.Office.WebExtension.Parameters.SliceSize]);
	}
});
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.GetDocumentCopyChunkAsync,
	requiredArguments: [
		{
			"name": Microsoft.Office.WebExtension.Parameters.SliceIndex,
			"types": ["number"]
		}
	],
	privateStateCallbacks: [
		{
			name: OSF.DDA.FileProperties.Handle,
			value: function (caller, stateInfo) { return stateInfo[OSF.DDA.FileProperties.Handle]; }
		},
		{
			name: OSF.DDA.FileProperties.SliceSize,
			value: function (caller, stateInfo) { return stateInfo[OSF.DDA.FileProperties.SliceSize]; }
		}
	],
	checkCallArgs: function (callArgs, caller, stateInfo) {
		var index=callArgs[Microsoft.Office.WebExtension.Parameters.SliceIndex];
		if (index < 0 || index >=caller.sliceCount) {
			throw OSF.DDA.ErrorCodeManager.errorCodes.ooeIndexOutOfRange;
		}
		callArgs[OSF.DDA.FileSliceOffset]=parseInt((index * stateInfo[OSF.DDA.FileProperties.SliceSize]).toString());
		return callArgs;
	},
	onSucceeded: function (sliceDescriptor, caller, callArgs) {
		var slice={};
		OSF.OUtil.defineEnumerableProperties(slice, {
			"data": {
				value: sliceDescriptor[Microsoft.Office.WebExtension.Parameters.Data]
			},
			"index": {
				value: callArgs[Microsoft.Office.WebExtension.Parameters.SliceIndex]
			},
			"size": {
				value: sliceDescriptor[OSF.DDA.FileProperties.SliceSize]
			}
		});
		return slice;
	}
});
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.ReleaseDocumentCopyAsync,
	privateStateCallbacks: [
		{
			name: OSF.DDA.FileProperties.Handle,
			value: function (caller, stateInfo) { return stateInfo[OSF.DDA.FileProperties.Handle]; }
		}
	]
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.FileType, {
	Text: "text",
	Pdf: "pdf"
});
Microsoft.Office.WebExtension.EventType={};
OSF.EventDispatch=function OSF_EventDispatch(eventTypes) {
	this._eventHandlers={};
	this._objectEventHandlers={};
	this._queuedEventsArgs={};
	if (eventTypes !=null) {
		for (var i=0; i < eventTypes.length; i++) {
			var eventType=eventTypes[i];
			var isObjectEvent=(eventType=="objectDeleted" || eventType=="objectSelectionChanged" || eventType=="objectDataChanged" || eventType=="contentControlAdded");
			if (!isObjectEvent)
				this._eventHandlers[eventType]=[];
			else
				this._objectEventHandlers[eventType]={};
			this._queuedEventsArgs[eventType]=[];
		}
	}
};
OSF.EventDispatch.prototype={
	getSupportedEvents: function OSF_EventDispatch$getSupportedEvents() {
		var events=[];
		for (var eventName in this._eventHandlers)
			events.push(eventName);
		for (var eventName in this._objectEventHandlers)
			events.push(eventName);
		return events;
	},
	supportsEvent: function OSF_EventDispatch$supportsEvent(event) {
		for (var eventName in this._eventHandlers) {
			if (event==eventName)
				return true;
		}
		for (var eventName in this._objectEventHandlers) {
			if (event==eventName)
				return true;
		}
		return false;
	},
	hasEventHandler: function OSF_EventDispatch$hasEventHandler(eventType, handler) {
		var handlers=this._eventHandlers[eventType];
		if (handlers && handlers.length > 0) {
			for (var i=0; i < handlers.length; i++) {
				if (handlers[i]===handler)
					return true;
			}
		}
		return false;
	},
	hasObjectEventHandler: function OSF_EventDispatch$hasObjectEventHandler(eventType, objectId, handler) {
		var handlers=this._objectEventHandlers[eventType];
		if (handlers !=null) {
			var _handlers=handlers[objectId];
			for (var i=0; _handlers !=null && i < _handlers.length; i++) {
				if (_handlers[i]===handler)
					return true;
			}
		}
		return false;
	},
	addEventHandler: function OSF_EventDispatch$addEventHandler(eventType, handler) {
		if (typeof handler !="function") {
			return false;
		}
		var handlers=this._eventHandlers[eventType];
		if (handlers && !this.hasEventHandler(eventType, handler)) {
			handlers.push(handler);
			return true;
		}
		else {
			return false;
		}
	},
	addObjectEventHandler: function OSF_EventDispatch$addObjectEventHandler(eventType, objectId, handler) {
		if (typeof handler !="function") {
			return false;
		}
		var handlers=this._objectEventHandlers[eventType];
		if (handlers && !this.hasObjectEventHandler(eventType, objectId, handler)) {
			if (handlers[objectId]==null)
				handlers[objectId]=[];
			handlers[objectId].push(handler);
			return true;
		}
		return false;
	},
	addEventHandlerAndFireQueuedEvent: function OSF_EventDispatch$addEventHandlerAndFireQueuedEvent(eventType, handler) {
		var handlers=this._eventHandlers[eventType];
		var isFirstHandler=handlers.length==0;
		var succeed=this.addEventHandler(eventType, handler);
		if (isFirstHandler && succeed) {
			this.fireQueuedEvent(eventType);
		}
		return succeed;
	},
	removeEventHandler: function OSF_EventDispatch$removeEventHandler(eventType, handler) {
		var handlers=this._eventHandlers[eventType];
		if (handlers && handlers.length > 0) {
			for (var index=0; index < handlers.length; index++) {
				if (handlers[index]===handler) {
					handlers.splice(index, 1);
					return true;
				}
			}
		}
		return false;
	},
	removeObjectEventHandler: function OSF_EventDispatch$removeObjectEventHandler(eventType, objectId, handler) {
		var handlers=this._objectEventHandlers[eventType];
		if (handlers !=null) {
			var _handlers=handlers[objectId];
			for (var i=0; _handlers !=null && i < _handlers.length; i++) {
				if (_handlers[i]===handler) {
					_handlers.splice(i, 1);
					return true;
				}
			}
		}
		return false;
	},
	clearEventHandlers: function OSF_EventDispatch$clearEventHandlers(eventType) {
		if (typeof this._eventHandlers[eventType] !="undefined" && this._eventHandlers[eventType].length > 0) {
			this._eventHandlers[eventType]=[];
			return true;
		}
		return false;
	},
	clearObjectEventHandlers: function OSF_EventDispatch$clearObjectEventHandlers(eventType, objectId) {
		if (this._objectEventHandlers[eventType] !=null && this._objectEventHandlers[eventType][objectId] !=null) {
			this._objectEventHandlers[eventType][objectId]=[];
			return true;
		}
		return false;
	},
	getEventHandlerCount: function OSF_EventDispatch$getEventHandlerCount(eventType) {
		return this._eventHandlers[eventType] !=undefined ? this._eventHandlers[eventType].length : -1;
	},
	getObjectEventHandlerCount: function OSF_EventDispatch$getObjectEventHandlerCount(eventType, objectId) {
		if (this._objectEventHandlers[eventType]==null || this._objectEventHandlers[eventType][objectId]==null)
			return 0;
		return this._objectEventHandlers[eventType][objectId].length;
	},
	fireEvent: function OSF_EventDispatch$fireEvent(eventArgs) {
		if (eventArgs.type==undefined)
			return false;
		var eventType=eventArgs.type;
		if (eventType && this._eventHandlers[eventType]) {
			var eventHandlers=this._eventHandlers[eventType];
			for (var i=0; i < eventHandlers.length; i++) {
				eventHandlers[i](eventArgs);
			}
			return true;
		}
		else {
			return false;
		}
	},
	fireObjectEvent: function OSF_EventDispatch$fireObjectEvent(objectId, eventArgs) {
		if (eventArgs.type==undefined)
			return false;
		var eventType=eventArgs.type;
		if (eventType && this._objectEventHandlers[eventType]) {
			var eventHandlers=this._objectEventHandlers[eventType];
			var _handlers=eventHandlers[objectId];
			if (_handlers !=null) {
				for (var i=0; i < _handlers.length; i++)
					_handlers[i](eventArgs);
				return true;
			}
		}
		return false;
	},
	fireOrQueueEvent: function OSF_EventDispatch$fireOrQueueEvent(eventArgs) {
		var eventType=eventArgs.type;
		if (eventType && this._eventHandlers[eventType]) {
			var eventHandlers=this._eventHandlers[eventType];
			var queuedEvents=this._queuedEventsArgs[eventType];
			if (eventHandlers.length==0) {
				queuedEvents.push(eventArgs);
			}
			else {
				this.fireEvent(eventArgs);
			}
			return true;
		}
		else {
			return false;
		}
	},
	fireQueuedEvent: function OSF_EventDispatch$queueEvent(eventType) {
		if (eventType && this._eventHandlers[eventType]) {
			var eventHandlers=this._eventHandlers[eventType];
			var queuedEvents=this._queuedEventsArgs[eventType];
			if (eventHandlers.length > 0) {
				var eventHandler=eventHandlers[0];
				while (queuedEvents.length > 0) {
					var eventArgs=queuedEvents.shift();
					eventHandler(eventArgs);
				}
				return true;
			}
		}
		return false;
	},
	clearQueuedEvent: function OSF_EventDispatch$clearQueuedEvent(eventType) {
		if (eventType && this._eventHandlers[eventType]) {
			var queuedEvents=this._queuedEventsArgs[eventType];
			if (queuedEvents) {
				this._queuedEventsArgs[eventType]=[];
			}
		}
	}
};
OSF.DDA.OMFactory=OSF.DDA.OMFactory || {};
OSF.DDA.OMFactory.manufactureEventArgs=function OSF_DDA_OMFactory$manufactureEventArgs(eventType, target, eventProperties) {
	var args;
	switch (eventType) {
		case Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged:
			args=new OSF.DDA.DocumentSelectionChangedEventArgs(target);
			break;
		case Microsoft.Office.WebExtension.EventType.BindingSelectionChanged:
			args=new OSF.DDA.BindingSelectionChangedEventArgs(this.manufactureBinding(eventProperties, target.document), eventProperties[OSF.DDA.PropertyDescriptors.Subset]);
			break;
		case Microsoft.Office.WebExtension.EventType.BindingDataChanged:
			args=new OSF.DDA.BindingDataChangedEventArgs(this.manufactureBinding(eventProperties, target.document));
			break;
		case Microsoft.Office.WebExtension.EventType.SettingsChanged:
			args=new OSF.DDA.SettingsChangedEventArgs(target);
			break;
		case Microsoft.Office.WebExtension.EventType.ActiveViewChanged:
			args=new OSF.DDA.ActiveViewChangedEventArgs(eventProperties);
			break;
		case Microsoft.Office.WebExtension.EventType.OfficeThemeChanged:
			args=new OSF.DDA.Theming.OfficeThemeChangedEventArgs(eventProperties);
			break;
		case Microsoft.Office.WebExtension.EventType.DocumentThemeChanged:
			args=new OSF.DDA.Theming.DocumentThemeChangedEventArgs(eventProperties);
			break;
		case Microsoft.Office.WebExtension.EventType.AppCommandInvoked:
			args=OSF.DDA.AppCommand.AppCommandInvokedEventArgs.create(eventProperties);
			break;
		case Microsoft.Office.WebExtension.EventType.ObjectDeleted:
		case Microsoft.Office.WebExtension.EventType.ObjectSelectionChanged:
		case Microsoft.Office.WebExtension.EventType.ObjectDataChanged:
		case Microsoft.Office.WebExtension.EventType.ContentControlAdded:
			args=new OSF.DDA.ObjectEventArgs(eventType, eventProperties[Microsoft.Office.WebExtension.Parameters.Id]);
			break;
		case Microsoft.Office.WebExtension.EventType.RichApiMessage:
			args=new OSF.DDA.RichApiMessageEventArgs(eventType, eventProperties);
			break;
		case Microsoft.Office.WebExtension.EventType.DataNodeInserted:
			args=new OSF.DDA.NodeInsertedEventArgs(this.manufactureDataNode(eventProperties[OSF.DDA.DataNodeEventProperties.NewNode]), eventProperties[OSF.DDA.DataNodeEventProperties.InUndoRedo]);
			break;
		case Microsoft.Office.WebExtension.EventType.DataNodeReplaced:
			args=new OSF.DDA.NodeReplacedEventArgs(this.manufactureDataNode(eventProperties[OSF.DDA.DataNodeEventProperties.OldNode]), this.manufactureDataNode(eventProperties[OSF.DDA.DataNodeEventProperties.NewNode]), eventProperties[OSF.DDA.DataNodeEventProperties.InUndoRedo]);
			break;
		case Microsoft.Office.WebExtension.EventType.DataNodeDeleted:
			args=new OSF.DDA.NodeDeletedEventArgs(this.manufactureDataNode(eventProperties[OSF.DDA.DataNodeEventProperties.OldNode]), this.manufactureDataNode(eventProperties[OSF.DDA.DataNodeEventProperties.NextSiblingNode]), eventProperties[OSF.DDA.DataNodeEventProperties.InUndoRedo]);
			break;
		case Microsoft.Office.WebExtension.EventType.TaskSelectionChanged:
			args=new OSF.DDA.TaskSelectionChangedEventArgs(target);
			break;
		case Microsoft.Office.WebExtension.EventType.ResourceSelectionChanged:
			args=new OSF.DDA.ResourceSelectionChangedEventArgs(target);
			break;
		case Microsoft.Office.WebExtension.EventType.ViewSelectionChanged:
			args=new OSF.DDA.ViewSelectionChangedEventArgs(target);
			break;
		case Microsoft.Office.WebExtension.EventType.DialogMessageReceived:
			args=new OSF.DDA.DialogEventArgs(eventProperties);
			break;
		case Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived:
			args=new OSF.DDA.DialogParentEventArgs(eventProperties);
			break;
		case Microsoft.Office.WebExtension.EventType.ItemChanged:
			if (OSF._OfficeAppFactory.getHostInfo()["hostType"]=="outlook") {
				args=new OSF.DDA.OlkItemSelectedChangedEventArgs(eventProperties);
				target.initialize(args["initialData"]);
				if (OSF._OfficeAppFactory.getHostInfo()["hostPlatform"]=="win32" || OSF._OfficeAppFactory.getHostInfo()["hostPlatform"]=="mac") {
					target.setCurrentItemNumber(args["itemNumber"].itemNumber);
				}
			}
			else {
				throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
			}
			break;
		case Microsoft.Office.WebExtension.EventType.RecipientsChanged:
			if (OSF._OfficeAppFactory.getHostInfo()["hostType"]=="outlook") {
				args=new OSF.DDA.OlkRecipientsChangedEventArgs(eventProperties);
			}
			else {
				throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
			}
			break;
		case Microsoft.Office.WebExtension.EventType.AppointmentTimeChanged:
			if (OSF._OfficeAppFactory.getHostInfo()["hostType"]=="outlook") {
				args=new OSF.DDA.OlkAppointmentTimeChangedEventArgs(eventProperties);
			}
			else {
				throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
			}
			break;
		case Microsoft.Office.WebExtension.EventType.RecurrenceChanged:
			if (OSF._OfficeAppFactory.getHostInfo()["hostType"]=="outlook") {
				args=new OSF.DDA.OlkRecurrenceChangedEventArgs(eventProperties);
			}
			else {
				throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
			}
			break;
		default:
			throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
	}
	return args;
};
OSF.DDA.AsyncMethodNames.addNames({
	AddHandlerAsync: "addHandlerAsync",
	RemoveHandlerAsync: "removeHandlerAsync"
});
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.AddHandlerAsync,
	requiredArguments: [{
			"name": Microsoft.Office.WebExtension.Parameters.EventType,
			"enum": Microsoft.Office.WebExtension.EventType,
			"verify": function (eventType, caller, eventDispatch) { return eventDispatch.supportsEvent(eventType); }
		},
		{
			"name": Microsoft.Office.WebExtension.Parameters.Handler,
			"types": ["function"]
		}
	],
	supportedOptions: [],
	privateStateCallbacks: []
});
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.RemoveHandlerAsync,
	requiredArguments: [
		{
			"name": Microsoft.Office.WebExtension.Parameters.EventType,
			"enum": Microsoft.Office.WebExtension.EventType,
			"verify": function (eventType, caller, eventDispatch) { return eventDispatch.supportsEvent(eventType); }
		}
	],
	supportedOptions: [
		{
			name: Microsoft.Office.WebExtension.Parameters.Handler,
			value: {
				"types": ["function", "object"],
				"defaultValue": null
			}
		}
	],
	privateStateCallbacks: []
});
OSF.DialogShownStatus={ hasDialogShown: false, isWindowDialog: false };
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, {
	DialogMessageReceivedEvent: "DialogMessageReceivedEvent"
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
	DialogMessageReceived: "dialogMessageReceived",
	DialogEventReceived: "dialogEventReceived"
});
OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors, {
	MessageType: "messageType",
	MessageContent: "messageContent"
});
OSF.DDA.DialogEventType={};
OSF.OUtil.augmentList(OSF.DDA.DialogEventType, {
	DialogClosed: "dialogClosed",
	NavigationFailed: "naviationFailed"
});
OSF.DDA.AsyncMethodNames.addNames({
	DisplayDialogAsync: "displayDialogAsync",
	CloseAsync: "close"
});
OSF.DDA.SyncMethodNames.addNames({
	MessageParent: "messageParent",
	AddMessageHandler: "addEventHandler",
	SendMessage: "sendMessage"
});
OSF.DDA.UI.ParentUI=function OSF_DDA_ParentUI() {
	var eventDispatch=new OSF.EventDispatch([
		Microsoft.Office.WebExtension.EventType.DialogMessageReceived,
		Microsoft.Office.WebExtension.EventType.DialogEventReceived,
		Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived
	]);
	var openDialogName=OSF.DDA.AsyncMethodNames.DisplayDialogAsync.displayName;
	var target=this;
	if (!target[openDialogName]) {
		OSF.OUtil.defineEnumerableProperty(target, openDialogName, {
			value: function () {
				var openDialog=OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.OpenDialog];
				openDialog(arguments, eventDispatch, target);
			}
		});
	}
	OSF.OUtil.finalizeProperties(this);
};
OSF.DDA.UI.ChildUI=function OSF_DDA_ChildUI(isPopupWindow) {
	var messageParentName=OSF.DDA.SyncMethodNames.MessageParent.displayName;
	var target=this;
	if (!target[messageParentName]) {
		OSF.OUtil.defineEnumerableProperty(target, messageParentName, {
			value: function () {
				var messageParent=OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.MessageParent];
				return messageParent(arguments, target);
			}
		});
	}
	var addEventHandler=OSF.DDA.SyncMethodNames.AddMessageHandler.displayName;
	if (!target[addEventHandler] && typeof OSF.DialogParentMessageEventDispatch !="undefined") {
		OSF.DDA.DispIdHost.addEventSupport(target, OSF.DialogParentMessageEventDispatch, isPopupWindow);
	}
	OSF.OUtil.finalizeProperties(this);
};
OSF.DialogHandler=function OSF_DialogHandler() { };
OSF.DDA.DialogEventArgs=function OSF_DDA_DialogEventArgs(message) {
	if (message[OSF.DDA.PropertyDescriptors.MessageType]==OSF.DialogMessageType.DialogMessageReceived) {
		OSF.OUtil.defineEnumerableProperties(this, {
			"type": {
				value: Microsoft.Office.WebExtension.EventType.DialogMessageReceived
			},
			"message": {
				value: message[OSF.DDA.PropertyDescriptors.MessageContent]
			}
		});
	}
	else {
		OSF.OUtil.defineEnumerableProperties(this, {
			"type": {
				value: Microsoft.Office.WebExtension.EventType.DialogEventReceived
			},
			"error": {
				value: message[OSF.DDA.PropertyDescriptors.MessageType]
			}
		});
	}
};
OSF.DDA.DialogParentEventArgs=function OSF_DDA_DialogParentEventArgs(message) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived
		},
		"message": {
			value: message[OSF.DDA.PropertyDescriptors.MessageContent]
		}
	});
};
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.DisplayDialogAsync,
	requiredArguments: [
		{
			"name": Microsoft.Office.WebExtension.Parameters.Url,
			"types": ["string"]
		}
	],
	supportedOptions: [
		{
			name: Microsoft.Office.WebExtension.Parameters.Width,
			value: {
				"types": ["number"],
				"defaultValue": 99
			}
		},
		{
			name: Microsoft.Office.WebExtension.Parameters.Height,
			value: {
				"types": ["number"],
				"defaultValue": 99
			}
		},
		{
			name: Microsoft.Office.WebExtension.Parameters.RequireHTTPs,
			value: {
				"types": ["boolean"],
				"defaultValue": true
			}
		},
		{
			name: Microsoft.Office.WebExtension.Parameters.DisplayInIframe,
			value: {
				"types": ["boolean"],
				"defaultValue": false
			}
		},
		{
			name: Microsoft.Office.WebExtension.Parameters.HideTitle,
			value: {
				"types": ["boolean"],
				"defaultValue": false
			}
		},
		{
			name: Microsoft.Office.WebExtension.Parameters.UseDeviceIndependentPixels,
			value: {
				"types": ["boolean"],
				"defaultValue": false
			}
		}
	],
	privateStateCallbacks: [],
	onSucceeded: function (args, caller, callArgs) {
		var targetId=args[Microsoft.Office.WebExtension.Parameters.Id];
		var eventDispatch=args[Microsoft.Office.WebExtension.Parameters.Data];
		var dialog=new OSF.DialogHandler();
		var closeDialog=OSF.DDA.AsyncMethodNames.CloseAsync.displayName;
		OSF.OUtil.defineEnumerableProperty(dialog, closeDialog, {
			value: function () {
				var closeDialogfunction=OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.CloseDialog];
				closeDialogfunction(arguments, targetId, eventDispatch, dialog);
			}
		});
		var addHandler=OSF.DDA.SyncMethodNames.AddMessageHandler.displayName;
		OSF.OUtil.defineEnumerableProperty(dialog, addHandler, {
			value: function () {
				var syncMethodCall=OSF.DDA.SyncMethodCalls[OSF.DDA.SyncMethodNames.AddMessageHandler.id];
				var callArgs=syncMethodCall.verifyAndExtractCall(arguments, dialog, eventDispatch);
				var eventType=callArgs[Microsoft.Office.WebExtension.Parameters.EventType];
				var handler=callArgs[Microsoft.Office.WebExtension.Parameters.Handler];
				return eventDispatch.addEventHandlerAndFireQueuedEvent(eventType, handler);
			}
		});
		var sendMessage=OSF.DDA.SyncMethodNames.SendMessage.displayName;
		OSF.OUtil.defineEnumerableProperty(dialog, sendMessage, {
			value: function () {
				var execute=OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.SendMessage];
				return execute(arguments, eventDispatch, dialog);
			}
		});
		return dialog;
	},
	checkCallArgs: function (callArgs, caller, stateInfo) {
		if (callArgs[Microsoft.Office.WebExtension.Parameters.Width] <=0) {
			callArgs[Microsoft.Office.WebExtension.Parameters.Width]=1;
		}
		if (!callArgs[Microsoft.Office.WebExtension.Parameters.UseDeviceIndependentPixels] && callArgs[Microsoft.Office.WebExtension.Parameters.Width] > 100) {
			callArgs[Microsoft.Office.WebExtension.Parameters.Width]=99;
		}
		if (callArgs[Microsoft.Office.WebExtension.Parameters.Height] <=0) {
			callArgs[Microsoft.Office.WebExtension.Parameters.Height]=1;
		}
		if (!callArgs[Microsoft.Office.WebExtension.Parameters.UseDeviceIndependentPixels] && callArgs[Microsoft.Office.WebExtension.Parameters.Height] > 100) {
			callArgs[Microsoft.Office.WebExtension.Parameters.Height]=99;
		}
		if (!callArgs[Microsoft.Office.WebExtension.Parameters.RequireHTTPs]) {
			callArgs[Microsoft.Office.WebExtension.Parameters.RequireHTTPs]=true;
		}
		return callArgs;
	}
});
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.CloseAsync,
	requiredArguments: [],
	supportedOptions: [],
	privateStateCallbacks: []
});
OSF.DDA.SyncMethodCalls.define({
	method: OSF.DDA.SyncMethodNames.MessageParent,
	requiredArguments: [
		{
			"name": Microsoft.Office.WebExtension.Parameters.MessageToParent,
			"types": ["string", "number", "boolean"]
		}
	],
	supportedOptions: []
});
OSF.DDA.SyncMethodCalls.define({
	method: OSF.DDA.SyncMethodNames.AddMessageHandler,
	requiredArguments: [
		{
			"name": Microsoft.Office.WebExtension.Parameters.EventType,
			"enum": Microsoft.Office.WebExtension.EventType,
			"verify": function (eventType, caller, eventDispatch) { return eventDispatch.supportsEvent(eventType); }
		},
		{
			"name": Microsoft.Office.WebExtension.Parameters.Handler,
			"types": ["function"]
		}
	],
	supportedOptions: []
});
OSF.DDA.SyncMethodCalls.define({
	method: OSF.DDA.SyncMethodNames.SendMessage,
	requiredArguments: [
		{
			"name": Microsoft.Office.WebExtension.Parameters.MessageContent,
			"types": ["string"]
		}
	],
	supportedOptions: [],
	privateStateCallbacks: []
});
OSF.DDA.SafeArray.Delegate.openDialog=function OSF_DDA_SafeArray_Delegate$OpenDialog(args) {
	try {
		if (args.onCalling) {
			args.onCalling();
		}
		var callback=OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent(true, args);
		OSF.ClientHostController.openDialog(args.dispId, args.targetId, function OSF_DDA_SafeArrayDelegate$RegisterEventAsync_OnEvent(eventDispId, payload) {
			if (args.onEvent) {
				args.onEvent(payload);
			}
			if (OSF.AppTelemetry) {
				OSF.AppTelemetry.onEventDone(args.dispId);
			}
		}, callback);
	}
	catch (ex) {
		OSF.DDA.SafeArray.Delegate._onException(ex, args);
	}
};
OSF.DDA.SafeArray.Delegate.closeDialog=function OSF_DDA_SafeArray_Delegate$CloseDialog(args) {
	if (args.onCalling) {
		args.onCalling();
	}
	var callback=OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent(false, args);
	try {
		OSF.ClientHostController.closeDialog(args.dispId, args.targetId, callback);
	}
	catch (ex) {
		OSF.DDA.SafeArray.Delegate._onException(ex, args);
	}
};
OSF.DDA.SafeArray.Delegate.messageParent=function OSF_DDA_SafeArray_Delegate$MessageParent(args) {
	try {
		if (args.onCalling) {
			args.onCalling();
		}
		var startTime=(new Date()).getTime();
		var result=OSF.ClientHostController.messageParent(args.hostCallArgs);
		if (args.onReceiving) {
			args.onReceiving();
		}
		if (OSF.AppTelemetry) {
			OSF.AppTelemetry.onMethodDone(args.dispId, args.hostCallArgs, Math.abs((new Date()).getTime() - startTime), result);
		}
		return result;
	}
	catch (ex) {
		return OSF.DDA.SafeArray.Delegate._onExceptionSyncMethod(ex);
	}
};
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.EventDispId.dispidDialogMessageReceivedEvent,
	fromHost: [
		{ name: OSF.DDA.EventDescriptors.DialogMessageReceivedEvent, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.EventDescriptors.DialogMessageReceivedEvent,
	fromHost: [
		{ name: OSF.DDA.PropertyDescriptors.MessageType, value: 0 },
		{ name: OSF.DDA.PropertyDescriptors.MessageContent, value: 1 }
	],
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.sendMessage=function OSF_DDA_SafeArray_Delegate$SendMessage(args) {
	try {
		if (args.onCalling) {
			args.onCalling();
		}
		var startTime=(new Date()).getTime();
		var result=OSF.ClientHostController.sendMessage(args.hostCallArgs);
		if (args.onReceiving) {
			args.onReceiving();
		}
		return result;
	}
	catch (ex) {
		return OSF.DDA.SafeArray.Delegate._onExceptionSyncMethod(ex);
	}
};
OSF.DDA.FilePropertiesDescriptor={
	Url: "Url"
};
OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors, {
	FilePropertiesDescriptor: "FilePropertiesDescriptor"
});
Microsoft.Office.WebExtension.FileProperties=function Microsoft_Office_WebExtension_FileProperties(filePropertiesDescriptor) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"url": {
			value: filePropertiesDescriptor[OSF.DDA.FilePropertiesDescriptor.Url]
		}
	});
};
OSF.DDA.AsyncMethodNames.addNames({ GetFilePropertiesAsync: "getFilePropertiesAsync" });
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.GetFilePropertiesAsync,
	fromHost: [
		{ name: OSF.DDA.PropertyDescriptors.FilePropertiesDescriptor, value: 0 }
	],
	requiredArguments: [],
	supportedOptions: [],
	onSucceeded: function (filePropertiesDescriptor, caller, callArgs) {
		return new Microsoft.Office.WebExtension.FileProperties(filePropertiesDescriptor);
	}
});
Microsoft.Office.WebExtension.GoToType={
	Binding: "binding",
	NamedItem: "namedItem",
	Slide: "slide",
	Index: "index"
};
Microsoft.Office.WebExtension.SelectionMode={
	Default: "default",
	Selected: "selected",
	None: "none"
};
Microsoft.Office.WebExtension.Index={
	First: "first",
	Last: "last",
	Next: "next",
	Previous: "previous"
};
OSF.DDA.AsyncMethodNames.addNames({ GoToByIdAsync: "goToByIdAsync" });
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.GoToByIdAsync,
	requiredArguments: [{
			"name": Microsoft.Office.WebExtension.Parameters.Id,
			"types": ["string", "number"]
		},
		{
			"name": Microsoft.Office.WebExtension.Parameters.GoToType,
			"enum": Microsoft.Office.WebExtension.GoToType
		}
	],
	supportedOptions: [
		{
			name: Microsoft.Office.WebExtension.Parameters.SelectionMode,
			value: {
				"enum": Microsoft.Office.WebExtension.SelectionMode,
				"defaultValue": Microsoft.Office.WebExtension.SelectionMode.Default
			}
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: Microsoft.Office.WebExtension.Parameters.GoToType,
	toHost: [
		{ name: Microsoft.Office.WebExtension.GoToType.Binding, value: 0 },
		{ name: Microsoft.Office.WebExtension.GoToType.NamedItem, value: 1 },
		{ name: Microsoft.Office.WebExtension.GoToType.Slide, value: 2 },
		{ name: Microsoft.Office.WebExtension.GoToType.Index, value: 3 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: Microsoft.Office.WebExtension.Parameters.SelectionMode,
	toHost: [
		{ name: Microsoft.Office.WebExtension.SelectionMode.Default, value: 0 },
		{ name: Microsoft.Office.WebExtension.SelectionMode.Selected, value: 1 },
		{ name: Microsoft.Office.WebExtension.SelectionMode.None, value: 2 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidNavigateToMethod,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.GoToType, value: 1 },
		{ name: Microsoft.Office.WebExtension.Parameters.SelectionMode, value: 2 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.PropertyDescriptors.FileProperties,
	fromHost: [
		{ name: OSF.DDA.FileProperties.Handle, value: 0 },
		{ name: OSF.DDA.FileProperties.FileSize, value: 1 }
	],
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.PropertyDescriptors.FileSliceProperties,
	fromHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: 0 },
		{ name: OSF.DDA.FileProperties.SliceSize, value: 1 }
	],
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: Microsoft.Office.WebExtension.Parameters.FileType,
	toHost: [
		{ name: Microsoft.Office.WebExtension.FileType.Text, value: 0 },
		{ name: Microsoft.Office.WebExtension.FileType.Compressed, value: 5 },
		{ name: Microsoft.Office.WebExtension.FileType.Pdf, value: 6 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidGetDocumentCopyMethod,
	toHost: [{ name: Microsoft.Office.WebExtension.Parameters.FileType, value: 0 }],
	fromHost: [
		{ name: OSF.DDA.PropertyDescriptors.FileProperties, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidGetDocumentCopyChunkMethod,
	toHost: [
		{ name: OSF.DDA.FileProperties.Handle, value: 0 },
		{ name: OSF.DDA.FileSliceOffset, value: 1 },
		{ name: OSF.DDA.FileProperties.SliceSize, value: 2 }
	],
	fromHost: [
		{ name: OSF.DDA.PropertyDescriptors.FileSliceProperties, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidReleaseDocumentCopyMethod,
	toHost: [{ name: OSF.DDA.FileProperties.Handle, value: 0 }]
});
OSF.DDA.AsyncMethodNames.addNames({
	GetSelectedDataAsync: "getSelectedDataAsync",
	SetSelectedDataAsync: "setSelectedDataAsync"
});
(function () {
	function processData(dataDescriptor, caller, callArgs) {
		var data=dataDescriptor[Microsoft.Office.WebExtension.Parameters.Data];
		if (OSF.DDA.TableDataProperties && data && (data[OSF.DDA.TableDataProperties.TableRows] !=undefined || data[OSF.DDA.TableDataProperties.TableHeaders] !=undefined)) {
			data=OSF.DDA.OMFactory.manufactureTableData(data);
		}
		data=OSF.DDA.DataCoercion.coerceData(data, callArgs[Microsoft.Office.WebExtension.Parameters.CoercionType]);
		return data==undefined ? null : data;
	}
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.GetSelectedDataAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.CoercionType,
				"enum": Microsoft.Office.WebExtension.CoercionType
			}
		],
		supportedOptions: [
			{
				name: Microsoft.Office.WebExtension.Parameters.ValueFormat,
				value: {
					"enum": Microsoft.Office.WebExtension.ValueFormat,
					"defaultValue": Microsoft.Office.WebExtension.ValueFormat.Unformatted
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.FilterType,
				value: {
					"enum": Microsoft.Office.WebExtension.FilterType,
					"defaultValue": Microsoft.Office.WebExtension.FilterType.All
				}
			}
		],
		privateStateCallbacks: [],
		onSucceeded: processData
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.SetSelectedDataAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.Data,
				"types": ["string", "object", "number", "boolean"]
			}
		],
		supportedOptions: [
			{
				name: Microsoft.Office.WebExtension.Parameters.CoercionType,
				value: {
					"enum": Microsoft.Office.WebExtension.CoercionType,
					"calculate": function (requiredArgs) {
						return OSF.DDA.DataCoercion.determineCoercionType(requiredArgs[Microsoft.Office.WebExtension.Parameters.Data]);
					}
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.ImageLeft,
				value: {
					"types": ["number", "boolean"],
					"defaultValue": false
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.ImageTop,
				value: {
					"types": ["number", "boolean"],
					"defaultValue": false
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.ImageWidth,
				value: {
					"types": ["number", "boolean"],
					"defaultValue": false
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.ImageHeight,
				value: {
					"types": ["number", "boolean"],
					"defaultValue": false
				}
			}
		],
		privateStateCallbacks: []
	});
})();
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidGetSelectedDataMethod,
	fromHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.CoercionType, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.ValueFormat, value: 1 },
		{ name: Microsoft.Office.WebExtension.Parameters.FilterType, value: 2 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidSetSelectedDataMethod,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.CoercionType, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: 1 },
		{ name: Microsoft.Office.WebExtension.Parameters.ImageLeft, value: 2 },
		{ name: Microsoft.Office.WebExtension.Parameters.ImageTop, value: 3 },
		{ name: Microsoft.Office.WebExtension.Parameters.ImageWidth, value: 4 },
		{ name: Microsoft.Office.WebExtension.Parameters.ImageHeight, value: 5 },
	]
});
OSF.DDA.SettingsManager={
	SerializedSettings: "serializedSettings",
	RefreshingSettings: "refreshingSettings",
	DateJSONPrefix: "Date(",
	DataJSONSuffix: ")",
	serializeSettings: function OSF_DDA_SettingsManager$serializeSettings(settingsCollection) {
		return OSF.OUtil.serializeSettings(settingsCollection);
	},
	deserializeSettings: function OSF_DDA_SettingsManager$deserializeSettings(serializedSettings) {
		return OSF.OUtil.deserializeSettings(serializedSettings);
	}
};
OSF.DDA.Settings=function OSF_DDA_Settings(settings) {
	settings=settings || {};
	var cacheSessionSettings=function (settings) {
		var osfSessionStorage=OSF.OUtil.getSessionStorage();
		if (osfSessionStorage) {
			var serializedSettings=OSF.DDA.SettingsManager.serializeSettings(settings);
			var storageSettings=JSON ? JSON.stringify(serializedSettings) : Sys.Serialization.JavaScriptSerializer.serialize(serializedSettings);
			osfSessionStorage.setItem(OSF._OfficeAppFactory.getCachedSessionSettingsKey(), storageSettings);
		}
	};
	OSF.OUtil.defineEnumerableProperties(this, {
		"get": {
			value: function OSF_DDA_Settings$get(name) {
				var e=Function._validateParams(arguments, [
					{ name: "name", type: String, mayBeNull: false }
				]);
				if (e)
					throw e;
				var setting=settings[name];
				return typeof (setting)==='undefined' ? null : setting;
			}
		},
		"set": {
			value: function OSF_DDA_Settings$set(name, value) {
				var e=Function._validateParams(arguments, [
					{ name: "name", type: String, mayBeNull: false },
					{ name: "value", mayBeNull: true }
				]);
				if (e)
					throw e;
				settings[name]=value;
				cacheSessionSettings(settings);
			}
		},
		"remove": {
			value: function OSF_DDA_Settings$remove(name) {
				var e=Function._validateParams(arguments, [
					{ name: "name", type: String, mayBeNull: false }
				]);
				if (e)
					throw e;
				delete settings[name];
				cacheSessionSettings(settings);
			}
		}
	});
	OSF.DDA.DispIdHost.addAsyncMethods(this, [OSF.DDA.AsyncMethodNames.SaveAsync], settings);
};
OSF.DDA.RefreshableSettings=function OSF_DDA_RefreshableSettings(settings) {
	OSF.DDA.RefreshableSettings.uber.constructor.call(this, settings);
	OSF.DDA.DispIdHost.addAsyncMethods(this, [OSF.DDA.AsyncMethodNames.RefreshAsync], settings);
	OSF.DDA.DispIdHost.addEventSupport(this, new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.SettingsChanged]));
};
OSF.OUtil.extend(OSF.DDA.RefreshableSettings, OSF.DDA.Settings);
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
	SettingsChanged: "settingsChanged"
});
OSF.DDA.SettingsChangedEventArgs=function OSF_DDA_SettingsChangedEventArgs(settingsInstance) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.EventType.SettingsChanged
		},
		"settings": {
			value: settingsInstance
		}
	});
};
OSF.DDA.AsyncMethodNames.addNames({
	RefreshAsync: "refreshAsync",
	SaveAsync: "saveAsync"
});
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.RefreshAsync,
	requiredArguments: [],
	supportedOptions: [],
	privateStateCallbacks: [
		{
			name: OSF.DDA.SettingsManager.RefreshingSettings,
			value: function getRefreshingSettings(settingsInstance, settingsCollection) {
				return settingsCollection;
			}
		}
	],
	onSucceeded: function deserializeSettings(serializedSettingsDescriptor, refreshingSettings, refreshingSettingsArgs) {
		var serializedSettings=serializedSettingsDescriptor[OSF.DDA.SettingsManager.SerializedSettings];
		var newSettings=OSF.DDA.SettingsManager.deserializeSettings(serializedSettings);
		var oldSettings=refreshingSettingsArgs[OSF.DDA.SettingsManager.RefreshingSettings];
		for (var setting in oldSettings) {
			refreshingSettings.remove(setting);
		}
		for (var setting in newSettings) {
			refreshingSettings.set(setting, newSettings[setting]);
		}
		return refreshingSettings;
	}
});
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.SaveAsync,
	requiredArguments: [],
	supportedOptions: [
		{
			name: Microsoft.Office.WebExtension.Parameters.OverwriteIfStale,
			value: {
				"types": ["boolean"],
				"defaultValue": true
			}
		}
	],
	privateStateCallbacks: [
		{
			name: OSF.DDA.SettingsManager.SerializedSettings,
			value: function serializeSettings(settingsInstance, settingsCollection) {
				return OSF.DDA.SettingsManager.serializeSettings(settingsCollection);
			}
		}
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidLoadSettingsMethod,
	fromHost: [
		{ name: OSF.DDA.SettingsManager.SerializedSettings, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidSaveSettingsMethod,
	toHost: [
		{ name: OSF.DDA.SettingsManager.SerializedSettings, value: OSF.DDA.SettingsManager.SerializedSettings },
		{ name: Microsoft.Office.WebExtension.Parameters.OverwriteIfStale, value: Microsoft.Office.WebExtension.Parameters.OverwriteIfStale }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({ type: OSF.DDA.EventDispId.dispidSettingsChangedEvent });
Microsoft.Office.WebExtension.BindingType={
	Table: "table",
	Text: "text",
	Matrix: "matrix"
};
OSF.DDA.BindingProperties={
	Id: "BindingId",
	Type: Microsoft.Office.WebExtension.Parameters.BindingType
};
OSF.OUtil.augmentList(OSF.DDA.ListDescriptors, { BindingList: "BindingList" });
OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors, {
	Subset: "subset",
	BindingProperties: "BindingProperties"
});
OSF.DDA.ListType.setListType(OSF.DDA.ListDescriptors.BindingList, OSF.DDA.PropertyDescriptors.BindingProperties);
OSF.DDA.BindingPromise=function OSF_DDA_BindingPromise(bindingId, errorCallback) {
	this._id=bindingId;
	OSF.OUtil.defineEnumerableProperty(this, "onFail", {
		get: function () {
			return errorCallback;
		},
		set: function (onError) {
			var t=typeof onError;
			if (t !="undefined" && t !="function") {
				throw OSF.OUtil.formatString(Strings.OfficeOM.L_CallbackNotAFunction, t);
			}
			errorCallback=onError;
		}
	});
};
OSF.DDA.BindingPromise.prototype={
	_fetch: function OSF_DDA_BindingPromise$_fetch(onComplete) {
		if (this.binding) {
			if (onComplete)
				onComplete(this.binding);
		}
		else {
			if (!this._binding) {
				var me=this;
				Microsoft.Office.WebExtension.context.document.bindings.getByIdAsync(this._id, function (asyncResult) {
					if (asyncResult.status==Microsoft.Office.WebExtension.AsyncResultStatus.Succeeded) {
						OSF.OUtil.defineEnumerableProperty(me, "binding", {
							value: asyncResult.value
						});
						if (onComplete)
							onComplete(me.binding);
					}
					else {
						if (me.onFail)
							me.onFail(asyncResult);
					}
				});
			}
		}
		return this;
	},
	getDataAsync: function OSF_DDA_BindingPromise$getDataAsync() {
		var args=arguments;
		this._fetch(function onComplete(binding) { binding.getDataAsync.apply(binding, args); });
		return this;
	},
	setDataAsync: function OSF_DDA_BindingPromise$setDataAsync() {
		var args=arguments;
		this._fetch(function onComplete(binding) { binding.setDataAsync.apply(binding, args); });
		return this;
	},
	addHandlerAsync: function OSF_DDA_BindingPromise$addHandlerAsync() {
		var args=arguments;
		this._fetch(function onComplete(binding) { binding.addHandlerAsync.apply(binding, args); });
		return this;
	},
	removeHandlerAsync: function OSF_DDA_BindingPromise$removeHandlerAsync() {
		var args=arguments;
		this._fetch(function onComplete(binding) { binding.removeHandlerAsync.apply(binding, args); });
		return this;
	}
};
OSF.DDA.BindingFacade=function OSF_DDA_BindingFacade(docInstance) {
	this._eventDispatches=[];
	OSF.OUtil.defineEnumerableProperty(this, "document", {
		value: docInstance
	});
	var am=OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(this, [
		am.AddFromSelectionAsync,
		am.AddFromNamedItemAsync,
		am.GetAllAsync,
		am.GetByIdAsync,
		am.ReleaseByIdAsync
	]);
};
OSF.DDA.UnknownBinding=function OSF_DDA_UknonwnBinding(id, docInstance) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"document": { value: docInstance },
		"id": { value: id }
	});
};
OSF.DDA.Binding=function OSF_DDA_Binding(id, docInstance) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"document": {
			value: docInstance
		},
		"id": {
			value: id
		}
	});
	var am=OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(this, [
		am.GetDataAsync,
		am.SetDataAsync
	]);
	var et=Microsoft.Office.WebExtension.EventType;
	var bindingEventDispatches=docInstance.bindings._eventDispatches;
	if (!bindingEventDispatches[id]) {
		bindingEventDispatches[id]=new OSF.EventDispatch([
			et.BindingSelectionChanged,
			et.BindingDataChanged
		]);
	}
	var eventDispatch=bindingEventDispatches[id];
	OSF.DDA.DispIdHost.addEventSupport(this, eventDispatch);
};
OSF.DDA.generateBindingId=function OSF_DDA$GenerateBindingId() {
	return "UnnamedBinding_"+OSF.OUtil.getUniqueId()+"_"+new Date().getTime();
};
OSF.DDA.OMFactory=OSF.DDA.OMFactory || {};
OSF.DDA.OMFactory.manufactureBinding=function OSF_DDA_OMFactory$manufactureBinding(bindingProperties, containingDocument) {
	var id=bindingProperties[OSF.DDA.BindingProperties.Id];
	var rows=bindingProperties[OSF.DDA.BindingProperties.RowCount];
	var cols=bindingProperties[OSF.DDA.BindingProperties.ColumnCount];
	var hasHeaders=bindingProperties[OSF.DDA.BindingProperties.HasHeaders];
	var binding;
	switch (bindingProperties[OSF.DDA.BindingProperties.Type]) {
		case Microsoft.Office.WebExtension.BindingType.Text:
			binding=new OSF.DDA.TextBinding(id, containingDocument);
			break;
		case Microsoft.Office.WebExtension.BindingType.Matrix:
			binding=new OSF.DDA.MatrixBinding(id, containingDocument, rows, cols);
			break;
		case Microsoft.Office.WebExtension.BindingType.Table:
			var isExcelApp=function () {
				return (OSF.DDA.ExcelDocument)
					&& (Microsoft.Office.WebExtension.context.document)
					&& (Microsoft.Office.WebExtension.context.document instanceof OSF.DDA.ExcelDocument);
			};
			var tableBindingObject;
			if (isExcelApp() && OSF.DDA.ExcelTableBinding) {
				tableBindingObject=OSF.DDA.ExcelTableBinding;
			}
			else {
				tableBindingObject=OSF.DDA.TableBinding;
			}
			binding=new tableBindingObject(id, containingDocument, rows, cols, hasHeaders);
			break;
		default:
			binding=new OSF.DDA.UnknownBinding(id, containingDocument);
	}
	return binding;
};
OSF.DDA.AsyncMethodNames.addNames({
	AddFromSelectionAsync: "addFromSelectionAsync",
	AddFromNamedItemAsync: "addFromNamedItemAsync",
	GetAllAsync: "getAllAsync",
	GetByIdAsync: "getByIdAsync",
	ReleaseByIdAsync: "releaseByIdAsync",
	GetDataAsync: "getDataAsync",
	SetDataAsync: "setDataAsync"
});
(function () {
	function processBinding(bindingDescriptor) {
		return OSF.DDA.OMFactory.manufactureBinding(bindingDescriptor, Microsoft.Office.WebExtension.context.document);
	}
	function getObjectId(obj) { return obj.id; }
	function processData(dataDescriptor, caller, callArgs) {
		var data=dataDescriptor[Microsoft.Office.WebExtension.Parameters.Data];
		if (OSF.DDA.TableDataProperties && data && (data[OSF.DDA.TableDataProperties.TableRows] !=undefined || data[OSF.DDA.TableDataProperties.TableHeaders] !=undefined)) {
			data=OSF.DDA.OMFactory.manufactureTableData(data);
		}
		data=OSF.DDA.DataCoercion.coerceData(data, callArgs[Microsoft.Office.WebExtension.Parameters.CoercionType]);
		return data==undefined ? null : data;
	}
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.AddFromSelectionAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.BindingType,
				"enum": Microsoft.Office.WebExtension.BindingType
			}
		],
		supportedOptions: [{
				name: Microsoft.Office.WebExtension.Parameters.Id,
				value: {
					"types": ["string"],
					"calculate": OSF.DDA.generateBindingId
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.Columns,
				value: {
					"types": ["object"],
					"defaultValue": null
				}
			}
		],
		privateStateCallbacks: [],
		onSucceeded: processBinding
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.AddFromNamedItemAsync,
		requiredArguments: [{
				"name": Microsoft.Office.WebExtension.Parameters.ItemName,
				"types": ["string"]
			},
			{
				"name": Microsoft.Office.WebExtension.Parameters.BindingType,
				"enum": Microsoft.Office.WebExtension.BindingType
			}
		],
		supportedOptions: [{
				name: Microsoft.Office.WebExtension.Parameters.Id,
				value: {
					"types": ["string"],
					"calculate": OSF.DDA.generateBindingId
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.Columns,
				value: {
					"types": ["object"],
					"defaultValue": null
				}
			}
		],
		privateStateCallbacks: [
			{
				name: Microsoft.Office.WebExtension.Parameters.FailOnCollision,
				value: function () { return true; }
			}
		],
		onSucceeded: processBinding
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.GetAllAsync,
		requiredArguments: [],
		supportedOptions: [],
		privateStateCallbacks: [],
		onSucceeded: function (response) { return OSF.OUtil.mapList(response[OSF.DDA.ListDescriptors.BindingList], processBinding); }
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.GetByIdAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.Id,
				"types": ["string"]
			}
		],
		supportedOptions: [],
		privateStateCallbacks: [],
		onSucceeded: processBinding
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.ReleaseByIdAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.Id,
				"types": ["string"]
			}
		],
		supportedOptions: [],
		privateStateCallbacks: [],
		onSucceeded: function (response, caller, callArgs) {
			var id=callArgs[Microsoft.Office.WebExtension.Parameters.Id];
			delete caller._eventDispatches[id];
		}
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.GetDataAsync,
		requiredArguments: [],
		supportedOptions: [{
				name: Microsoft.Office.WebExtension.Parameters.CoercionType,
				value: {
					"enum": Microsoft.Office.WebExtension.CoercionType,
					"calculate": function (requiredArgs, binding) { return OSF.DDA.DataCoercion.getCoercionDefaultForBinding(binding.type); }
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.ValueFormat,
				value: {
					"enum": Microsoft.Office.WebExtension.ValueFormat,
					"defaultValue": Microsoft.Office.WebExtension.ValueFormat.Unformatted
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.FilterType,
				value: {
					"enum": Microsoft.Office.WebExtension.FilterType,
					"defaultValue": Microsoft.Office.WebExtension.FilterType.All
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.Rows,
				value: {
					"types": ["object", "string"],
					"defaultValue": null
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.Columns,
				value: {
					"types": ["object"],
					"defaultValue": null
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.StartRow,
				value: {
					"types": ["number"],
					"defaultValue": 0
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.StartColumn,
				value: {
					"types": ["number"],
					"defaultValue": 0
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.RowCount,
				value: {
					"types": ["number"],
					"defaultValue": 0
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.ColumnCount,
				value: {
					"types": ["number"],
					"defaultValue": 0
				}
			}
		],
		checkCallArgs: function (callArgs, caller, stateInfo) {
			if (callArgs[Microsoft.Office.WebExtension.Parameters.StartRow]==0 &&
				callArgs[Microsoft.Office.WebExtension.Parameters.StartColumn]==0 &&
				callArgs[Microsoft.Office.WebExtension.Parameters.RowCount]==0 &&
				callArgs[Microsoft.Office.WebExtension.Parameters.ColumnCount]==0) {
				delete callArgs[Microsoft.Office.WebExtension.Parameters.StartRow];
				delete callArgs[Microsoft.Office.WebExtension.Parameters.StartColumn];
				delete callArgs[Microsoft.Office.WebExtension.Parameters.RowCount];
				delete callArgs[Microsoft.Office.WebExtension.Parameters.ColumnCount];
			}
			if (callArgs[Microsoft.Office.WebExtension.Parameters.CoercionType] !=OSF.DDA.DataCoercion.getCoercionDefaultForBinding(caller.type) &&
				(callArgs[Microsoft.Office.WebExtension.Parameters.StartRow] ||
					callArgs[Microsoft.Office.WebExtension.Parameters.StartColumn] ||
					callArgs[Microsoft.Office.WebExtension.Parameters.RowCount] ||
					callArgs[Microsoft.Office.WebExtension.Parameters.ColumnCount])) {
				throw OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotMatchBinding;
			}
			return callArgs;
		},
		privateStateCallbacks: [
			{
				name: Microsoft.Office.WebExtension.Parameters.Id,
				value: getObjectId
			}
		],
		onSucceeded: processData
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.SetDataAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.Data,
				"types": ["string", "object", "number", "boolean"]
			}
		],
		supportedOptions: [{
				name: Microsoft.Office.WebExtension.Parameters.CoercionType,
				value: {
					"enum": Microsoft.Office.WebExtension.CoercionType,
					"calculate": function (requiredArgs) { return OSF.DDA.DataCoercion.determineCoercionType(requiredArgs[Microsoft.Office.WebExtension.Parameters.Data]); }
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.Rows,
				value: {
					"types": ["object", "string"],
					"defaultValue": null
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.Columns,
				value: {
					"types": ["object"],
					"defaultValue": null
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.StartRow,
				value: {
					"types": ["number"],
					"defaultValue": 0
				}
			},
			{
				name: Microsoft.Office.WebExtension.Parameters.StartColumn,
				value: {
					"types": ["number"],
					"defaultValue": 0
				}
			}
		],
		checkCallArgs: function (callArgs, caller, stateInfo) {
			if (callArgs[Microsoft.Office.WebExtension.Parameters.StartRow]==0 &&
				callArgs[Microsoft.Office.WebExtension.Parameters.StartColumn]==0) {
				delete callArgs[Microsoft.Office.WebExtension.Parameters.StartRow];
				delete callArgs[Microsoft.Office.WebExtension.Parameters.StartColumn];
			}
			if (callArgs[Microsoft.Office.WebExtension.Parameters.CoercionType] !=OSF.DDA.DataCoercion.getCoercionDefaultForBinding(caller.type) &&
				(callArgs[Microsoft.Office.WebExtension.Parameters.StartRow] ||
					callArgs[Microsoft.Office.WebExtension.Parameters.StartColumn])) {
				throw OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotMatchBinding;
			}
			return callArgs;
		},
		privateStateCallbacks: [
			{
				name: Microsoft.Office.WebExtension.Parameters.Id,
				value: getObjectId
			}
		]
	});
})();
OSF.OUtil.augmentList(OSF.DDA.BindingProperties, {
	RowCount: "BindingRowCount",
	ColumnCount: "BindingColumnCount",
	HasHeaders: "HasHeaders"
});
OSF.DDA.MatrixBinding=function OSF_DDA_MatrixBinding(id, docInstance, rows, cols) {
	OSF.DDA.MatrixBinding.uber.constructor.call(this, id, docInstance);
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.BindingType.Matrix
		},
		"rowCount": {
			value: rows ? rows : 0
		},
		"columnCount": {
			value: cols ? cols : 0
		}
	});
};
OSF.OUtil.extend(OSF.DDA.MatrixBinding, OSF.DDA.Binding);
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.PropertyDescriptors.BindingProperties,
	fromHost: [
		{ name: OSF.DDA.BindingProperties.Id, value: 0 },
		{ name: OSF.DDA.BindingProperties.Type, value: 1 },
		{ name: OSF.DDA.SafeArray.UniqueArguments.BindingSpecificData, value: 2 }
	],
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: Microsoft.Office.WebExtension.Parameters.BindingType,
	toHost: [
		{ name: Microsoft.Office.WebExtension.BindingType.Text, value: 0 },
		{ name: Microsoft.Office.WebExtension.BindingType.Matrix, value: 1 },
		{ name: Microsoft.Office.WebExtension.BindingType.Table, value: 2 }
	],
	invertible: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidAddBindingFromSelectionMethod,
	fromHost: [
		{ name: OSF.DDA.PropertyDescriptors.BindingProperties, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.BindingType, value: 1 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidAddBindingFromNamedItemMethod,
	fromHost: [
		{ name: OSF.DDA.PropertyDescriptors.BindingProperties, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.ItemName, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 1 },
		{ name: Microsoft.Office.WebExtension.Parameters.BindingType, value: 2 },
		{ name: Microsoft.Office.WebExtension.Parameters.FailOnCollision, value: 3 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidReleaseBindingMethod,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidGetBindingMethod,
	fromHost: [
		{ name: OSF.DDA.PropertyDescriptors.BindingProperties, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidGetAllBindingsMethod,
	fromHost: [
		{ name: OSF.DDA.ListDescriptors.BindingList, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidGetBindingDataMethod,
	fromHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.CoercionType, value: 1 },
		{ name: Microsoft.Office.WebExtension.Parameters.ValueFormat, value: 2 },
		{ name: Microsoft.Office.WebExtension.Parameters.FilterType, value: 3 },
		{ name: OSF.DDA.PropertyDescriptors.Subset, value: 4 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidSetBindingDataMethod,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.CoercionType, value: 1 },
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: 2 },
		{ name: OSF.DDA.SafeArray.UniqueArguments.Offset, value: 3 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.SafeArray.UniqueArguments.BindingSpecificData,
	fromHost: [
		{ name: OSF.DDA.BindingProperties.RowCount, value: 0 },
		{ name: OSF.DDA.BindingProperties.ColumnCount, value: 1 },
		{ name: OSF.DDA.BindingProperties.HasHeaders, value: 2 }
	],
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.PropertyDescriptors.Subset,
	toHost: [
		{ name: OSF.DDA.SafeArray.UniqueArguments.Offset, value: 0 },
		{ name: OSF.DDA.SafeArray.UniqueArguments.Run, value: 1 }
	],
	canonical: true,
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.SafeArray.UniqueArguments.Offset,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.StartRow, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.StartColumn, value: 1 }
	],
	canonical: true,
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.SafeArray.UniqueArguments.Run,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.RowCount, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.ColumnCount, value: 1 }
	],
	canonical: true,
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidAddRowsMethod,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: 1 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidAddColumnsMethod,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: 1 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidClearAllRowsMethod,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 }
	]
});
Microsoft.Office.WebExtension.TableData=function Microsoft_Office_WebExtension_TableData(rows, headers) {
	function fixData(data) {
		if (data==null || data==undefined) {
			return null;
		}
		try {
			for (var dim=OSF.DDA.DataCoercion.findArrayDimensionality(data, 2); dim < 2; dim++) {
				data=[data];
			}
			return data;
		}
		catch (ex) {
		}
	}
	;
	OSF.OUtil.defineEnumerableProperties(this, {
		"headers": {
			get: function () { return headers; },
			set: function (value) {
				headers=fixData(value);
			}
		},
		"rows": {
			get: function () { return rows; },
			set: function (value) {
				rows=(value==null || (OSF.OUtil.isArray(value) && (value.length==0))) ?
					[] :
					fixData(value);
			}
		}
	});
	this.headers=headers;
	this.rows=rows;
};
OSF.DDA.OMFactory=OSF.DDA.OMFactory || {};
OSF.DDA.OMFactory.manufactureTableData=function OSF_DDA_OMFactory$manufactureTableData(tableDataProperties) {
	return new Microsoft.Office.WebExtension.TableData(tableDataProperties[OSF.DDA.TableDataProperties.TableRows], tableDataProperties[OSF.DDA.TableDataProperties.TableHeaders]);
};
OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors, { TableDataProperties: "TableDataProperties" });
OSF.OUtil.augmentList(OSF.DDA.BindingProperties, {
	RowCount: "BindingRowCount",
	ColumnCount: "BindingColumnCount",
	HasHeaders: "HasHeaders"
});
OSF.DDA.TableDataProperties={
	TableRows: "TableRows",
	TableHeaders: "TableHeaders"
};
OSF.DDA.TableBinding=function OSF_DDA_TableBinding(id, docInstance, rows, cols, hasHeaders) {
	OSF.DDA.TableBinding.uber.constructor.call(this, id, docInstance);
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.BindingType.Table
		},
		"rowCount": {
			value: rows ? rows : 0
		},
		"columnCount": {
			value: cols ? cols : 0
		},
		"hasHeaders": {
			value: hasHeaders ? hasHeaders : false
		}
	});
	var am=OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(this, [
		am.AddRowsAsync,
		am.AddColumnsAsync,
		am.DeleteAllDataValuesAsync
	]);
};
OSF.OUtil.extend(OSF.DDA.TableBinding, OSF.DDA.Binding);
OSF.DDA.AsyncMethodNames.addNames({
	AddRowsAsync: "addRowsAsync",
	AddColumnsAsync: "addColumnsAsync",
	DeleteAllDataValuesAsync: "deleteAllDataValuesAsync"
});
(function () {
	function getObjectId(obj) { return obj.id; }
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.AddRowsAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.Data,
				"types": ["object"]
			}
		],
		supportedOptions: [],
		privateStateCallbacks: [
			{
				name: Microsoft.Office.WebExtension.Parameters.Id,
				value: getObjectId
			}
		]
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.AddColumnsAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.Data,
				"types": ["object"]
			}
		],
		supportedOptions: [],
		privateStateCallbacks: [
			{
				name: Microsoft.Office.WebExtension.Parameters.Id,
				value: getObjectId
			}
		]
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.DeleteAllDataValuesAsync,
		requiredArguments: [],
		supportedOptions: [],
		privateStateCallbacks: [
			{
				name: Microsoft.Office.WebExtension.Parameters.Id,
				value: getObjectId
			}
		]
	});
})();
OSF.DDA.TextBinding=function OSF_DDA_TextBinding(id, docInstance) {
	OSF.DDA.TextBinding.uber.constructor.call(this, id, docInstance);
	OSF.OUtil.defineEnumerableProperty(this, "type", {
		value: Microsoft.Office.WebExtension.BindingType.Text
	});
};
OSF.OUtil.extend(OSF.DDA.TextBinding, OSF.DDA.Binding);
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, { DocumentSelectionChanged: "documentSelectionChanged" });
OSF.DDA.DocumentSelectionChangedEventArgs=function OSF_DDA_DocumentSelectionChangedEventArgs(docInstance) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged
		},
		"document": {
			value: docInstance
		}
	});
};
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, { ObjectDeleted: "objectDeleted" });
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, { ObjectSelectionChanged: "objectSelectionChanged" });
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, { ObjectDataChanged: "objectDataChanged" });
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, { ContentControlAdded: "contentControlAdded" });
OSF.DDA.ObjectEventArgs=function OSF_DDA_ObjectEventArgs(eventType, object) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": { value: eventType },
		"object": { value: object }
	});
};
OSF.DDA.SafeArray.Delegate.ParameterMap.define({ type: OSF.DDA.EventDispId.dispidDocumentSelectionChangedEvent });
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.EventDispId.dispidObjectDeletedEvent,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 }
	],
	fromHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: OSF.DDA.SafeArray.Delegate.ParameterMap.sourceData }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.EventDispId.dispidObjectSelectionChangedEvent,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 }
	],
	fromHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: OSF.DDA.SafeArray.Delegate.ParameterMap.sourceData }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.EventDispId.dispidObjectDataChangedEvent,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 }
	],
	fromHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: OSF.DDA.SafeArray.Delegate.ParameterMap.sourceData }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.EventDispId.dispidContentControlAddedEvent,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 }
	],
	fromHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: OSF.DDA.SafeArray.Delegate.ParameterMap.sourceData }
	]
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
	BindingSelectionChanged: "bindingSelectionChanged",
	BindingDataChanged: "bindingDataChanged"
});
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, { BindingSelectionChangedEvent: "BindingSelectionChangedEvent" });
OSF.DDA.BindingSelectionChangedEventArgs=function OSF_DDA_BindingSelectionChangedEventArgs(bindingInstance, subset) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.EventType.BindingSelectionChanged
		},
		"binding": {
			value: bindingInstance
		}
	});
	for (var prop in subset) {
		OSF.OUtil.defineEnumerableProperty(this, prop, {
			value: subset[prop]
		});
	}
};
OSF.DDA.BindingDataChangedEventArgs=function OSF_DDA_BindingDataChangedEventArgs(bindingInstance) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.EventType.BindingDataChanged
		},
		"binding": {
			value: bindingInstance
		}
	});
};
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.EventDescriptors.BindingSelectionChangedEvent,
	fromHost: [
		{ name: OSF.DDA.PropertyDescriptors.BindingProperties, value: 0 },
		{ name: OSF.DDA.PropertyDescriptors.Subset, value: 1 }
	],
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.EventDispId.dispidBindingSelectionChangedEvent,
	fromHost: [
		{ name: OSF.DDA.EventDescriptors.BindingSelectionChangedEvent, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.EventDispId.dispidBindingDataChangedEvent,
	fromHost: [{ name: OSF.DDA.PropertyDescriptors.BindingProperties, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }]
});
Microsoft.Office.WebExtension.CoercionType={
	Text: "text",
	Matrix: "matrix",
	Table: "table"
};
OSF.DDA.DataCoercion=(function OSF_DDA_DataCoercion() {
	return {
		findArrayDimensionality: function OSF_DDA_DataCoercion$findArrayDimensionality(obj) {
			if (OSF.OUtil.isArray(obj)) {
				var dim=0;
				for (var index=0; index < obj.length; index++) {
					dim=Math.max(dim, OSF.DDA.DataCoercion.findArrayDimensionality(obj[index]));
				}
				return dim+1;
			}
			else {
				return 0;
			}
		},
		getCoercionDefaultForBinding: function OSF_DDA_DataCoercion$getCoercionDefaultForBinding(bindingType) {
			switch (bindingType) {
				case Microsoft.Office.WebExtension.BindingType.Matrix: return Microsoft.Office.WebExtension.CoercionType.Matrix;
				case Microsoft.Office.WebExtension.BindingType.Table: return Microsoft.Office.WebExtension.CoercionType.Table;
				case Microsoft.Office.WebExtension.BindingType.Text:
				default:
					return Microsoft.Office.WebExtension.CoercionType.Text;
			}
		},
		getBindingDefaultForCoercion: function OSF_DDA_DataCoercion$getBindingDefaultForCoercion(coercionType) {
			switch (coercionType) {
				case Microsoft.Office.WebExtension.CoercionType.Matrix: return Microsoft.Office.WebExtension.BindingType.Matrix;
				case Microsoft.Office.WebExtension.CoercionType.Table: return Microsoft.Office.WebExtension.BindingType.Table;
				case Microsoft.Office.WebExtension.CoercionType.Text:
				case Microsoft.Office.WebExtension.CoercionType.Html:
				case Microsoft.Office.WebExtension.CoercionType.Ooxml:
				default:
					return Microsoft.Office.WebExtension.BindingType.Text;
			}
		},
		determineCoercionType: function OSF_DDA_DataCoercion$determineCoercionType(data) {
			if (data==null || data==undefined)
				return null;
			var sourceType=null;
			var runtimeType=typeof data;
			if (data.rows !==undefined) {
				sourceType=Microsoft.Office.WebExtension.CoercionType.Table;
			}
			else if (OSF.OUtil.isArray(data)) {
				sourceType=Microsoft.Office.WebExtension.CoercionType.Matrix;
			}
			else if (runtimeType=="string" || runtimeType=="number" || runtimeType=="boolean" || OSF.OUtil.isDate(data)) {
				sourceType=Microsoft.Office.WebExtension.CoercionType.Text;
			}
			else {
				throw OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedDataObject;
			}
			return sourceType;
		},
		coerceData: function OSF_DDA_DataCoercion$coerceData(data, destinationType, sourceType) {
			sourceType=sourceType || OSF.DDA.DataCoercion.determineCoercionType(data);
			if (sourceType && sourceType !=destinationType) {
				OSF.OUtil.writeProfilerMark(OSF.InternalPerfMarker.DataCoercionBegin);
				data=OSF.DDA.DataCoercion._coerceDataFromTable(destinationType, OSF.DDA.DataCoercion._coerceDataToTable(data, sourceType));
				OSF.OUtil.writeProfilerMark(OSF.InternalPerfMarker.DataCoercionEnd);
			}
			return data;
		},
		_matrixToText: function OSF_DDA_DataCoercion$_matrixToText(matrix) {
			if (matrix.length==1 && matrix[0].length==1)
				return ""+matrix[0][0];
			var val="";
			for (var i=0; i < matrix.length; i++) {
				val+=matrix[i].join("\t")+"\n";
			}
			return val.substring(0, val.length - 1);
		},
		_textToMatrix: function OSF_DDA_DataCoercion$_textToMatrix(text) {
			var ret=text.split("\n");
			for (var i=0; i < ret.length; i++)
				ret[i]=ret[i].split("\t");
			return ret;
		},
		_tableToText: function OSF_DDA_DataCoercion$_tableToText(table) {
			var headers="";
			if (table.headers !=null) {
				headers=OSF.DDA.DataCoercion._matrixToText([table.headers])+"\n";
			}
			var rows=OSF.DDA.DataCoercion._matrixToText(table.rows);
			if (rows=="") {
				headers=headers.substring(0, headers.length - 1);
			}
			return headers+rows;
		},
		_tableToMatrix: function OSF_DDA_DataCoercion$_tableToMatrix(table) {
			var matrix=table.rows;
			if (table.headers !=null) {
				matrix.unshift(table.headers);
			}
			return matrix;
		},
		_coerceDataFromTable: function OSF_DDA_DataCoercion$_coerceDataFromTable(coercionType, table) {
			var value;
			switch (coercionType) {
				case Microsoft.Office.WebExtension.CoercionType.Table:
					value=table;
					break;
				case Microsoft.Office.WebExtension.CoercionType.Matrix:
					value=OSF.DDA.DataCoercion._tableToMatrix(table);
					break;
				case Microsoft.Office.WebExtension.CoercionType.SlideRange:
					value=null;
					if (OSF.DDA.OMFactory.manufactureSlideRange) {
						value=OSF.DDA.OMFactory.manufactureSlideRange(OSF.DDA.DataCoercion._tableToText(table));
					}
					if (value==null) {
						value=OSF.DDA.DataCoercion._tableToText(table);
					}
					break;
				case Microsoft.Office.WebExtension.CoercionType.Text:
				case Microsoft.Office.WebExtension.CoercionType.Html:
				case Microsoft.Office.WebExtension.CoercionType.Ooxml:
				default:
					value=OSF.DDA.DataCoercion._tableToText(table);
					break;
			}
			return value;
		},
		_coerceDataToTable: function OSF_DDA_DataCoercion$_coerceDataToTable(data, sourceType) {
			if (sourceType==undefined) {
				sourceType=OSF.DDA.DataCoercion.determineCoercionType(data);
			}
			var value;
			switch (sourceType) {
				case Microsoft.Office.WebExtension.CoercionType.Table:
					value=data;
					break;
				case Microsoft.Office.WebExtension.CoercionType.Matrix:
					value=new Microsoft.Office.WebExtension.TableData(data);
					break;
				case Microsoft.Office.WebExtension.CoercionType.Text:
				case Microsoft.Office.WebExtension.CoercionType.Html:
				case Microsoft.Office.WebExtension.CoercionType.Ooxml:
				default:
					value=new Microsoft.Office.WebExtension.TableData(OSF.DDA.DataCoercion._textToMatrix(data));
					break;
			}
			return value;
		}
	};
})();
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: Microsoft.Office.WebExtension.Parameters.CoercionType,
	toHost: [
		{ name: Microsoft.Office.WebExtension.CoercionType.Text, value: 0 },
		{ name: Microsoft.Office.WebExtension.CoercionType.Matrix, value: 1 },
		{ name: Microsoft.Office.WebExtension.CoercionType.Table, value: 2 }
	]
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.CoercionType, { Html: "html" });
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: Microsoft.Office.WebExtension.Parameters.CoercionType,
	toHost: [
		{ name: Microsoft.Office.WebExtension.CoercionType.Html, value: 3 }
	]
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.CoercionType, { Ooxml: "ooxml" });
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: Microsoft.Office.WebExtension.Parameters.CoercionType,
	toHost: [
		{ name: Microsoft.Office.WebExtension.CoercionType.Ooxml, value: 4 }
	]
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.CoercionType, { OoxmlPackage: "ooxmlPackage" });
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: Microsoft.Office.WebExtension.Parameters.CoercionType,
	toHost: [
		{ name: Microsoft.Office.WebExtension.CoercionType.OoxmlPackage, value: 5 }
	]
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.CoercionType, { PdfFile: "pdf" });
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: Microsoft.Office.WebExtension.Parameters.CoercionType,
	toHost: [
		{ name: Microsoft.Office.WebExtension.CoercionType.PdfFile, value: 6 }
	]
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.FilterType, { OnlyVisible: "onlyVisible" });
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: Microsoft.Office.WebExtension.Parameters.FilterType,
	toHost: [{ name: Microsoft.Office.WebExtension.FilterType.OnlyVisible, value: 1 }]
});
OSF.DDA.DataPartProperties={
	Id: Microsoft.Office.WebExtension.Parameters.Id,
	BuiltIn: "DataPartBuiltIn"
};
OSF.DDA.DataNodeProperties={
	Handle: "DataNodeHandle",
	BaseName: "DataNodeBaseName",
	NamespaceUri: "DataNodeNamespaceUri",
	NodeType: "DataNodeType"
};
OSF.DDA.DataNodeEventProperties={
	OldNode: "OldNode",
	NewNode: "NewNode",
	NextSiblingNode: "NextSiblingNode",
	InUndoRedo: "InUndoRedo"
};
OSF.OUtil.augmentList(OSF.DDA.PropertyDescriptors, {
	DataPartProperties: "DataPartProperties",
	DataNodeProperties: "DataNodeProperties"
});
OSF.OUtil.augmentList(OSF.DDA.ListDescriptors, {
	DataPartList: "DataPartList",
	DataNodeList: "DataNodeList"
});
OSF.DDA.ListType.setListType(OSF.DDA.ListDescriptors.DataPartList, OSF.DDA.PropertyDescriptors.DataPartProperties);
OSF.DDA.ListType.setListType(OSF.DDA.ListDescriptors.DataNodeList, OSF.DDA.PropertyDescriptors.DataNodeProperties);
OSF.OUtil.augmentList(OSF.DDA.EventDescriptors, {
	DataNodeInsertedEvent: "DataNodeInsertedEvent",
	DataNodeReplacedEvent: "DataNodeReplacedEvent",
	DataNodeDeletedEvent: "DataNodeDeletedEvent"
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, {
	DataNodeDeleted: "nodeDeleted",
	DataNodeInserted: "nodeInserted",
	DataNodeReplaced: "nodeReplaced",
	NodeDeleted: "nodeDeleted",
	NodeInserted: "nodeInserted",
	NodeReplaced: "nodeReplaced"
});
OSF.DDA.CustomXmlParts=function OSF_DDA_CustomXmlParts() {
	this._eventDispatches=[];
	var am=OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(this, [
		am.AddDataPartAsync,
		am.GetDataPartByIdAsync,
		am.GetDataPartsByNameSpaceAsync
	]);
};
OSF.DDA.CustomXmlPart=function OSF_DDA_CustomXmlPart(customXmlParts, id, builtIn) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"builtIn": {
			value: builtIn
		},
		"id": {
			value: id
		},
		"namespaceManager": {
			value: new OSF.DDA.CustomXmlPrefixMappings(id)
		}
	});
	var am=OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(this, [
		am.DeleteDataPartAsync,
		am.GetPartNodesAsync,
		am.GetPartXmlAsync
	]);
	var customXmlPartEventDispatches=customXmlParts._eventDispatches;
	var dispatch=customXmlPartEventDispatches[id];
	if (!dispatch) {
		var et=Microsoft.Office.WebExtension.EventType;
		dispatch=new OSF.EventDispatch([
			et.DataNodeDeleted,
			et.DataNodeInserted,
			et.DataNodeReplaced
		]);
		customXmlPartEventDispatches[id]=dispatch;
	}
	OSF.DDA.DispIdHost.addEventSupport(this, dispatch);
};
OSF.DDA.CustomXmlPrefixMappings=function OSF_DDA_CustomXmlPrefixMappings(partId) {
	var am=OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(this, [
		am.AddDataPartNamespaceAsync,
		am.GetDataPartNamespaceAsync,
		am.GetDataPartPrefixAsync
	], partId);
};
OSF.DDA.CustomXmlNode=function OSF_DDA_CustomXmlNode(handle, nodeType, ns, baseName) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"baseName": {
			value: baseName
		},
		"namespaceUri": {
			value: ns
		},
		"nodeType": {
			value: nodeType
		}
	});
	var am=OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(this, [
		am.GetRelativeNodesAsync,
		am.GetNodeValueAsync,
		am.GetNodeXmlAsync,
		am.SetNodeValueAsync,
		am.SetNodeXmlAsync,
		am.GetNodeTextAsync,
		am.SetNodeTextAsync
	], handle);
};
OSF.DDA.NodeInsertedEventArgs=function OSF_DDA_NodeInsertedEventArgs(newNode, inUndoRedo) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.EventType.DataNodeInserted
		},
		"newNode": {
			value: newNode
		},
		"inUndoRedo": {
			value: inUndoRedo
		}
	});
};
OSF.DDA.NodeReplacedEventArgs=function OSF_DDA_NodeReplacedEventArgs(oldNode, newNode, inUndoRedo) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.EventType.DataNodeReplaced
		},
		"oldNode": {
			value: oldNode
		},
		"newNode": {
			value: newNode
		},
		"inUndoRedo": {
			value: inUndoRedo
		}
	});
};
OSF.DDA.NodeDeletedEventArgs=function OSF_DDA_NodeDeletedEventArgs(oldNode, oldNextSibling, inUndoRedo) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.EventType.DataNodeDeleted
		},
		"oldNode": {
			value: oldNode
		},
		"oldNextSibling": {
			value: oldNextSibling
		},
		"inUndoRedo": {
			value: inUndoRedo
		}
	});
};
OSF.DDA.OMFactory=OSF.DDA.OMFactory || {};
OSF.DDA.OMFactory.manufactureDataNode=function OSF_DDA_OMFactory$manufactureDataNode(nodeProperties) {
	if (nodeProperties) {
		return new OSF.DDA.CustomXmlNode(nodeProperties[OSF.DDA.DataNodeProperties.Handle], nodeProperties[OSF.DDA.DataNodeProperties.NodeType], nodeProperties[OSF.DDA.DataNodeProperties.NamespaceUri], nodeProperties[OSF.DDA.DataNodeProperties.BaseName]);
	}
};
OSF.DDA.OMFactory.manufactureDataPart=function OSF_DDA_OMFactory$manufactureDataPart(partProperties, containingCustomXmlParts) {
	return new OSF.DDA.CustomXmlPart(containingCustomXmlParts, partProperties[OSF.DDA.DataPartProperties.Id], partProperties[OSF.DDA.DataPartProperties.BuiltIn]);
};
OSF.DDA.AsyncMethodNames.addNames({
	AddDataPartAsync: "addAsync",
	GetDataPartByIdAsync: "getByIdAsync",
	GetDataPartsByNameSpaceAsync: "getByNamespaceAsync",
	DeleteDataPartAsync: "deleteAsync",
	GetPartNodesAsync: "getNodesAsync",
	GetPartXmlAsync: "getXmlAsync",
	AddDataPartNamespaceAsync: "addNamespaceAsync",
	GetDataPartNamespaceAsync: "getNamespaceAsync",
	GetDataPartPrefixAsync: "getPrefixAsync",
	GetRelativeNodesAsync: "getNodesAsync",
	GetNodeValueAsync: "getNodeValueAsync",
	GetNodeXmlAsync: "getXmlAsync",
	SetNodeValueAsync: "setNodeValueAsync",
	SetNodeXmlAsync: "setXmlAsync",
	GetNodeTextAsync: "getTextAsync",
	SetNodeTextAsync: "setTextAsync"
});
(function () {
	function processDataPart(dataPartDescriptor) {
		return OSF.DDA.OMFactory.manufactureDataPart(dataPartDescriptor, Microsoft.Office.WebExtension.context.document.customXmlParts);
	}
	function processDataNode(dataNodeDescriptor) {
		return OSF.DDA.OMFactory.manufactureDataNode(dataNodeDescriptor);
	}
	function processData(dataDescriptor, caller, callArgs) {
		var data=dataDescriptor[Microsoft.Office.WebExtension.Parameters.Data];
		return data==undefined ? null : data;
	}
	function getObjectId(obj) { return obj.id; }
	function getPartId(part, partId) { return partId; }
	;
	function getNodeHandle(node, nodeHandle) { return nodeHandle; }
	;
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.AddDataPartAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.Xml,
				"types": ["string"]
			}
		],
		supportedOptions: [],
		privateStateCallbacks: [],
		onSucceeded: processDataPart
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.GetDataPartByIdAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.Id,
				"types": ["string"]
			}
		],
		supportedOptions: [],
		privateStateCallbacks: [],
		onSucceeded: processDataPart
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.GetDataPartsByNameSpaceAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.Namespace,
				"types": ["string"]
			}
		],
		supportedOptions: [],
		privateStateCallbacks: [],
		onSucceeded: function (response) { return OSF.OUtil.mapList(response[OSF.DDA.ListDescriptors.DataPartList], processDataPart); }
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.DeleteDataPartAsync,
		requiredArguments: [],
		supportedOptions: [],
		privateStateCallbacks: [
			{
				name: OSF.DDA.DataPartProperties.Id,
				value: getObjectId
			}
		]
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.GetPartNodesAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.XPath,
				"types": ["string"]
			}
		],
		supportedOptions: [],
		privateStateCallbacks: [
			{
				name: OSF.DDA.DataPartProperties.Id,
				value: getObjectId
			}
		],
		onSucceeded: function (response) { return OSF.OUtil.mapList(response[OSF.DDA.ListDescriptors.DataNodeList], processDataNode); }
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.GetPartXmlAsync,
		requiredArguments: [],
		supportedOptions: [],
		privateStateCallbacks: [
			{
				name: OSF.DDA.DataPartProperties.Id,
				value: getObjectId
			}
		],
		onSucceeded: processData
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.AddDataPartNamespaceAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.Prefix,
				"types": ["string"]
			},
			{
				"name": Microsoft.Office.WebExtension.Parameters.Namespace,
				"types": ["string"]
			}
		],
		supportedOptions: [],
		privateStateCallbacks: [
			{
				name: OSF.DDA.DataPartProperties.Id,
				value: getPartId
			}
		]
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.GetDataPartNamespaceAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.Prefix,
				"types": ["string"]
			}
		],
		supportedOptions: [],
		privateStateCallbacks: [
			{
				name: OSF.DDA.DataPartProperties.Id,
				value: getPartId
			}
		],
		onSucceeded: processData
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.GetDataPartPrefixAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.Namespace,
				"types": ["string"]
			}
		],
		supportedOptions: [],
		privateStateCallbacks: [
			{
				name: OSF.DDA.DataPartProperties.Id,
				value: getPartId
			}
		],
		onSucceeded: processData
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.GetRelativeNodesAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.XPath,
				"types": ["string"]
			}
		],
		supportedOptions: [],
		privateStateCallbacks: [
			{
				name: OSF.DDA.DataNodeProperties.Handle,
				value: getNodeHandle
			}
		],
		onSucceeded: function (response) { return OSF.OUtil.mapList(response[OSF.DDA.ListDescriptors.DataNodeList], processDataNode); }
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.GetNodeValueAsync,
		requiredArguments: [],
		supportedOptions: [],
		privateStateCallbacks: [
			{
				name: OSF.DDA.DataNodeProperties.Handle,
				value: getNodeHandle
			}
		],
		onSucceeded: processData
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.GetNodeXmlAsync,
		requiredArguments: [],
		supportedOptions: [],
		privateStateCallbacks: [
			{
				name: OSF.DDA.DataNodeProperties.Handle,
				value: getNodeHandle
			}
		],
		onSucceeded: processData
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.SetNodeValueAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.Data,
				"types": ["string"]
			}
		],
		supportedOptions: [],
		privateStateCallbacks: [
			{
				name: OSF.DDA.DataNodeProperties.Handle,
				value: getNodeHandle
			}
		]
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.SetNodeXmlAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.Xml,
				"types": ["string"]
			}
		],
		supportedOptions: [],
		privateStateCallbacks: [
			{
				name: OSF.DDA.DataNodeProperties.Handle,
				value: getNodeHandle
			}
		]
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.GetNodeTextAsync,
		requiredArguments: [],
		supportedOptions: [],
		privateStateCallbacks: [
			{
				name: OSF.DDA.DataNodeProperties.Handle,
				value: getNodeHandle
			}
		],
		onSucceeded: processData
	});
	OSF.DDA.AsyncMethodCalls.define({
		method: OSF.DDA.AsyncMethodNames.SetNodeTextAsync,
		requiredArguments: [
			{
				"name": Microsoft.Office.WebExtension.Parameters.Text,
				"types": ["string"]
			}
		],
		supportedOptions: [],
		privateStateCallbacks: [
			{
				name: OSF.DDA.DataNodeProperties.Handle,
				value: getNodeHandle
			}
		]
	});
})();
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.PropertyDescriptors.DataPartProperties,
	fromHost: [
		{ name: OSF.DDA.DataPartProperties.Id, value: 0 },
		{ name: OSF.DDA.DataPartProperties.BuiltIn, value: 1 }
	],
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.PropertyDescriptors.DataNodeProperties,
	fromHost: [
		{ name: OSF.DDA.DataNodeProperties.Handle, value: 0 },
		{ name: OSF.DDA.DataNodeProperties.BaseName, value: 1 },
		{ name: OSF.DDA.DataNodeProperties.NamespaceUri, value: 2 },
		{ name: OSF.DDA.DataNodeProperties.NodeType, value: 3 }
	],
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.EventDescriptors.DataNodeInsertedEvent,
	fromHost: [
		{ name: OSF.DDA.DataNodeEventProperties.InUndoRedo, value: 0 },
		{ name: OSF.DDA.DataNodeEventProperties.NewNode, value: 1 }
	],
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.EventDescriptors.DataNodeReplacedEvent,
	fromHost: [
		{ name: OSF.DDA.DataNodeEventProperties.InUndoRedo, value: 0 },
		{ name: OSF.DDA.DataNodeEventProperties.OldNode, value: 1 },
		{ name: OSF.DDA.DataNodeEventProperties.NewNode, value: 2 }
	],
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.EventDescriptors.DataNodeDeletedEvent,
	fromHost: [
		{ name: OSF.DDA.DataNodeEventProperties.InUndoRedo, value: 0 },
		{ name: OSF.DDA.DataNodeEventProperties.OldNode, value: 1 },
		{ name: OSF.DDA.DataNodeEventProperties.NextSiblingNode, value: 2 }
	],
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.DataNodeEventProperties.OldNode,
	fromHost: [
		{ name: OSF.DDA.PropertyDescriptors.DataNodeProperties, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.DataNodeEventProperties.NewNode,
	fromHost: [
		{ name: OSF.DDA.PropertyDescriptors.DataNodeProperties, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.DataNodeEventProperties.NextSiblingNode,
	fromHost: [
		{ name: OSF.DDA.PropertyDescriptors.DataNodeProperties, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidAddDataPartMethod,
	fromHost: [
		{ name: OSF.DDA.PropertyDescriptors.DataPartProperties, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Xml, value: 0 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidGetDataPartByIdMethod,
	fromHost: [
		{ name: OSF.DDA.PropertyDescriptors.DataPartProperties, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidGetDataPartsByNamespaceMethod,
	fromHost: [
		{ name: OSF.DDA.ListDescriptors.DataPartList, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Namespace, value: 0 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidGetDataPartXmlMethod,
	fromHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidGetDataPartNodesMethod,
	fromHost: [
		{ name: OSF.DDA.ListDescriptors.DataNodeList, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.XPath, value: 1 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidDeleteDataPartMethod,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Id, value: 0 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidGetDataNodeValueMethod,
	fromHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	toHost: [
		{ name: OSF.DDA.DataNodeProperties.Handle, value: 0 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidGetDataNodeXmlMethod,
	fromHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	toHost: [
		{ name: OSF.DDA.DataNodeProperties.Handle, value: 0 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidGetDataNodesMethod,
	fromHost: [
		{ name: OSF.DDA.ListDescriptors.DataNodeList, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	toHost: [
		{ name: OSF.DDA.DataNodeProperties.Handle, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.XPath, value: 1 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidSetDataNodeValueMethod,
	toHost: [
		{ name: OSF.DDA.DataNodeProperties.Handle, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: 1 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidSetDataNodeXmlMethod,
	toHost: [
		{ name: OSF.DDA.DataNodeProperties.Handle, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.Xml, value: 1 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidAddDataNamespaceMethod,
	toHost: [
		{ name: OSF.DDA.DataPartProperties.Id, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.Prefix, value: 1 },
		{ name: Microsoft.Office.WebExtension.Parameters.Namespace, value: 2 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidGetDataUriByPrefixMethod,
	fromHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	toHost: [
		{ name: OSF.DDA.DataPartProperties.Id, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.Prefix, value: 1 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidGetDataPrefixByUriMethod,
	fromHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	toHost: [
		{ name: OSF.DDA.DataPartProperties.Id, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.Namespace, value: 1 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidGetDataNodeTextMethod,
	fromHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	],
	toHost: [
		{ name: OSF.DDA.DataNodeProperties.Handle, value: 0 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidSetDataNodeTextMethod,
	toHost: [
		{ name: OSF.DDA.DataNodeProperties.Handle, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.Text, value: 1 }
	]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.EventDispId.dispidDataNodeAddedEvent,
	fromHost: [{ name: OSF.DDA.EventDescriptors.DataNodeInsertedEvent, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.EventDispId.dispidDataNodeReplacedEvent,
	fromHost: [{ name: OSF.DDA.EventDescriptors.DataNodeReplacedEvent, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.EventDispId.dispidDataNodeDeletedEvent,
	fromHost: [{ name: OSF.DDA.EventDescriptors.DataNodeDeletedEvent, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }]
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.PropertyDescriptors.FilePropertiesDescriptor,
	fromHost: [
		{ name: OSF.DDA.FilePropertiesDescriptor.Url, value: 0 }
	],
	isComplexType: true
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidGetFilePropertiesMethod,
	fromHost: [
		{ name: OSF.DDA.PropertyDescriptors.FilePropertiesDescriptor, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	]
});
OSF.DDA.AsyncMethodNames.addNames({
	ExecuteRichApiRequestAsync: "executeRichApiRequestAsync"
});
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.ExecuteRichApiRequestAsync,
	requiredArguments: [
		{
			name: Microsoft.Office.WebExtension.Parameters.Data,
			types: ["object"]
		}
	],
	supportedOptions: []
});
OSF.OUtil.setNamespace("RichApi", OSF.DDA);
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidExecuteRichApiRequestMethod,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: 0 }
	],
	fromHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	]
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.CoercionType, { Image: "image" });
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: Microsoft.Office.WebExtension.Parameters.CoercionType,
	toHost: [
		{ name: Microsoft.Office.WebExtension.CoercionType.Image, value: 8 }
	]
});
OSF.DDA.AsyncMethodNames.addNames({ GetAccessTokenAsync: "getAccessTokenAsync" });
OSF.DDA.Auth=function OSF_DDA_Auth() {
};
OSF.DDA.AsyncMethodCalls.define({
	method: OSF.DDA.AsyncMethodNames.GetAccessTokenAsync,
	requiredArguments: [],
	supportedOptions: [
		{
			name: Microsoft.Office.WebExtension.Parameters.ForceConsent,
			value: {
				"types": ["boolean"],
				"defaultValue": false
			}
		},
		{
			name: Microsoft.Office.WebExtension.Parameters.ForceAddAccount,
			value: {
				"types": ["boolean"],
				"defaultValue": false
			}
		},
		{
			name: Microsoft.Office.WebExtension.Parameters.AuthChallenge,
			value: {
				"types": ["string"],
				"defaultValue": ""
			}
		}
	],
	onSucceeded: function (dataDescriptor, caller, callArgs) {
		var data=dataDescriptor[Microsoft.Office.WebExtension.Parameters.Data];
		return data;
	}
});
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
	type: OSF.DDA.MethodDispId.dispidGetAccessTokenMethod,
	toHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.ForceConsent, value: 0 },
		{ name: Microsoft.Office.WebExtension.Parameters.ForceAddAccount, value: 1 },
		{ name: Microsoft.Office.WebExtension.Parameters.AuthChallenge, value: 2 }
	],
	fromHost: [
		{ name: Microsoft.Office.WebExtension.Parameters.Data, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
	]
});
OSF.DDA.WordDocument=function OSF_DDA_WordDocument(officeAppContext, settings) {
	OSF.DDA.WordDocument.uber.constructor.call(this, officeAppContext, new OSF.DDA.BindingFacade(this), settings);
	OSF.DDA.DispIdHost.addAsyncMethods(this, [
		OSF.DDA.AsyncMethodNames.GoToByIdAsync,
		OSF.DDA.AsyncMethodNames.GetDocumentCopyAsync,
		OSF.DDA.AsyncMethodNames.GetFilePropertiesAsync
	]);
	OSF.OUtil.defineEnumerableProperty(this, "customXmlParts", {
		value: new OSF.DDA.CustomXmlParts()
	});
	OSF.OUtil.finalizeProperties(this);
};
OSF.OUtil.extend(OSF.DDA.WordDocument, OSF.DDA.JsomDocument);
OSF.InitializationHelper.prototype.loadAppSpecificScriptAndCreateOM=function OSF_InitializationHelper$loadAppSpecificScriptAndCreateOM(appContext, appReady, basePath) {
	OSF.DDA.ErrorCodeManager.initializeErrorMessages(Strings.OfficeOM);
	appContext.doc=new OSF.DDA.WordDocument(appContext, this._initializeSettings(appContext, true));
	OSF.DDA.DispIdHost.addAsyncMethods(OSF.DDA.RichApi, [OSF.DDA.AsyncMethodNames.ExecuteRichApiRequestAsync]);
	appReady();
};

var __extends=(this && this.__extends) || function (d, b) {
	for (var p in b) if (b.hasOwnProperty(p)) d[p]=b[p];
	function __() { this.constructor=d; }
	d.prototype=b===null ? Object.create(b) : (__.prototype=b.prototype, new __());
};
var OfficeExtension;
(function (OfficeExtension) {
	var Action=(function () {
		function Action(actionInfo, isWriteOperation, isRestrictedResourceAccess) {
			this.m_actionInfo=actionInfo;
			this.m_isWriteOperation=isWriteOperation;
			this.m_isRestrictedResourceAccess=isRestrictedResourceAccess;
		}
		Object.defineProperty(Action.prototype, "actionInfo", {
			get: function () {
				return this.m_actionInfo;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Action.prototype, "isWriteOperation", {
			get: function () {
				return this.m_isWriteOperation;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Action.prototype, "isRestrictedResourceAccess", {
			get: function () {
				return this.m_isRestrictedResourceAccess;
			},
			enumerable: true,
			configurable: true
		});
		return Action;
	}());
	OfficeExtension.Action=Action;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var TraceMarkerActionResultHandler=(function () {
		function TraceMarkerActionResultHandler(callback) {
			this.m_callback=callback;
		}
		TraceMarkerActionResultHandler.prototype._handleResult=function (value) {
			if (this.m_callback) {
				this.m_callback();
			}
		};
		return TraceMarkerActionResultHandler;
	}());
	var ActionFactory=(function () {
		function ActionFactory() {
		}
		ActionFactory.createSetPropertyAction=function (context, parent, propertyName, value) {
			OfficeExtension.Utility.validateObjectPath(parent);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 4,
				Name: propertyName,
				ObjectPathId: parent._objectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			var args=[value];
			var referencedArgumentObjectPaths=OfficeExtension.Utility.setMethodArguments(context, actionInfo.ArgumentInfo, args);
			OfficeExtension.Utility.validateReferencedObjectPaths(referencedArgumentObjectPaths);
			context._pendingRequest.ensureInstantiateObjectPath(parent._objectPath);
			context._pendingRequest.ensureInstantiateObjectPaths(referencedArgumentObjectPaths);
			var ret=new OfficeExtension.Action(actionInfo, true, false);
			context._pendingRequest.addAction(ret);
			context._pendingRequest.addReferencedObjectPath(parent._objectPath);
			context._pendingRequest.addReferencedObjectPaths(referencedArgumentObjectPaths);
			ret.referencedObjectPath=parent._objectPath;
			ret.referencedArgumentObjectPaths=referencedArgumentObjectPaths;
			return ret;
		};
		ActionFactory.createMethodAction=function (context, parent, methodName, operationType, args, isRestrictedResourceAccess) {
			OfficeExtension.Utility.validateObjectPath(parent);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 3,
				Name: methodName,
				ObjectPathId: parent._objectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			var referencedArgumentObjectPaths=OfficeExtension.Utility.setMethodArguments(context, actionInfo.ArgumentInfo, args);
			OfficeExtension.Utility.validateReferencedObjectPaths(referencedArgumentObjectPaths);
			context._pendingRequest.ensureInstantiateObjectPath(parent._objectPath);
			context._pendingRequest.ensureInstantiateObjectPaths(referencedArgumentObjectPaths);
			var isWriteOperation=operationType !=1;
			var ret=new OfficeExtension.Action(actionInfo, isWriteOperation, isRestrictedResourceAccess);
			context._pendingRequest.addAction(ret);
			context._pendingRequest.addReferencedObjectPath(parent._objectPath);
			context._pendingRequest.addReferencedObjectPaths(referencedArgumentObjectPaths);
			ret.referencedObjectPath=parent._objectPath;
			ret.referencedArgumentObjectPaths=referencedArgumentObjectPaths;
			return ret;
		};
		ActionFactory.createQueryAction=function (context, parent, queryOption) {
			OfficeExtension.Utility.validateObjectPath(parent);
			context._pendingRequest.ensureInstantiateObjectPath(parent._objectPath);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 2,
				Name: "",
				ObjectPathId: parent._objectPath.objectPathInfo.Id,
			};
			actionInfo.QueryInfo=queryOption;
			var ret=new OfficeExtension.Action(actionInfo, false, false);
			context._pendingRequest.addAction(ret);
			context._pendingRequest.addReferencedObjectPath(parent._objectPath);
			ret.referencedObjectPath=parent._objectPath;
			return ret;
		};
		ActionFactory.createRecursiveQueryAction=function (context, parent, query) {
			OfficeExtension.Utility.validateObjectPath(parent);
			context._pendingRequest.ensureInstantiateObjectPath(parent._objectPath);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 6,
				Name: "",
				ObjectPathId: parent._objectPath.objectPathInfo.Id,
				RecursiveQueryInfo: query
			};
			var ret=new OfficeExtension.Action(actionInfo, false, false);
			context._pendingRequest.addAction(ret);
			context._pendingRequest.addReferencedObjectPath(parent._objectPath);
			ret.referencedObjectPath=parent._objectPath;
			return ret;
		};
		ActionFactory.createQueryAsJsonAction=function (context, parent, queryOption) {
			OfficeExtension.Utility.validateObjectPath(parent);
			context._pendingRequest.ensureInstantiateObjectPath(parent._objectPath);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 7,
				Name: "",
				ObjectPathId: parent._objectPath.objectPathInfo.Id,
			};
			actionInfo.QueryInfo=queryOption;
			var ret=new OfficeExtension.Action(actionInfo, false, false);
			context._pendingRequest.addAction(ret);
			context._pendingRequest.addReferencedObjectPath(parent._objectPath);
			ret.referencedObjectPath=parent._objectPath;
			return ret;
		};
		ActionFactory.createEnsureUnchangedAction=function (context, parent, objectState) {
			OfficeExtension.Utility.validateObjectPath(parent);
			context._pendingRequest.ensureInstantiateObjectPath(parent._objectPath);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 8,
				Name: "",
				ObjectPathId: parent._objectPath.objectPathInfo.Id,
				ObjectState: objectState
			};
			var ret=new OfficeExtension.Action(actionInfo, false, false);
			context._pendingRequest.addAction(ret);
			context._pendingRequest.addReferencedObjectPath(parent._objectPath);
			ret.referencedObjectPath=parent._objectPath;
			return ret;
		};
		ActionFactory.createUpdateAction=function (context, parent, objectState) {
			OfficeExtension.Utility.validateObjectPath(parent);
			context._pendingRequest.ensureInstantiateObjectPath(parent._objectPath);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 9,
				Name: "",
				ObjectPathId: parent._objectPath.objectPathInfo.Id,
				ObjectState: objectState
			};
			var ret=new OfficeExtension.Action(actionInfo, true, false);
			context._pendingRequest.addAction(ret);
			context._pendingRequest.addReferencedObjectPath(parent._objectPath);
			ret.referencedObjectPath=parent._objectPath;
			return ret;
		};
		ActionFactory.createInstantiateAction=function (context, obj) {
			OfficeExtension.Utility.validateObjectPath(obj);
			context._pendingRequest.ensureInstantiateObjectPath(obj._objectPath.parentObjectPath);
			context._pendingRequest.ensureInstantiateObjectPaths(obj._objectPath.argumentObjectPaths);
			var actionInfo={
				Id: context._nextId(),
				ActionType: 1,
				Name: "",
				ObjectPathId: obj._objectPath.objectPathInfo.Id
			};
			var ret=new OfficeExtension.Action(actionInfo, false, false);
			ret.referencedObjectPath=obj._objectPath;
			context._pendingRequest.addAction(ret);
			context._pendingRequest.addReferencedObjectPath(obj._objectPath);
			context._pendingRequest.addActionResultHandler(ret, new OfficeExtension.InstantiateActionResultHandler(obj));
			return ret;
		};
		ActionFactory.createTraceAction=function (context, message, addTraceMessage) {
			var actionInfo={
				Id: context._nextId(),
				ActionType: 5,
				Name: "Trace",
				ObjectPathId: 0
			};
			var ret=new OfficeExtension.Action(actionInfo, false, false);
			context._pendingRequest.addAction(ret);
			if (addTraceMessage) {
				context._pendingRequest.addTrace(actionInfo.Id, message);
			}
			return ret;
		};
		ActionFactory.createTraceMarkerForCallback=function (context, callback) {
			var action=ActionFactory.createTraceAction(context, null, false);
			context._pendingRequest.addActionResultHandler(action, new TraceMarkerActionResultHandler(callback));
		};
		return ActionFactory;
	}());
	OfficeExtension.ActionFactory=ActionFactory;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ClientObject=(function () {
		function ClientObject(context, objectPath) {
			OfficeExtension.Utility.checkArgumentNull(context, "context");
			this.m_context=context;
			this.m_objectPath=objectPath;
			if (this.m_objectPath) {
				if (!context._processingResult) {
					OfficeExtension.ActionFactory.createInstantiateAction(context, this);
					if ((context._autoCleanup) && (this._KeepReference)) {
						context.trackedObjects._autoAdd(this);
					}
				}
			}
		}
		Object.defineProperty(ClientObject.prototype, "context", {
			get: function () {
				return this.m_context;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientObject.prototype, "_objectPath", {
			get: function () {
				return this.m_objectPath;
			},
			set: function (value) {
				this.m_objectPath=value;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientObject.prototype, "isNull", {
			get: function () {
				OfficeExtension.Utility.throwIfNotLoaded("isNull", this._isNull, null, this._isNull);
				return this._isNull;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientObject.prototype, "isNullObject", {
			get: function () {
				OfficeExtension.Utility.throwIfNotLoaded("isNullObject", this._isNull, null, this._isNull);
				return this._isNull;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientObject.prototype, "_isNull", {
			get: function () {
				return this.m_isNull;
			},
			set: function (value) {
				this.m_isNull=value;
				if (value && this.m_objectPath) {
					this.m_objectPath._updateAsNullObject();
				}
			},
			enumerable: true,
			configurable: true
		});
		ClientObject.prototype._handleResult=function (value) {
			this._isNull=OfficeExtension.Utility.isNullOrUndefined(value);
			this.context.trackedObjects._autoTrackIfNecessaryWhenHandleObjectResultValue(this, value);
		};
		ClientObject.prototype._handleIdResult=function (value) {
			this._isNull=OfficeExtension.Utility.isNullOrUndefined(value);
			OfficeExtension.Utility.fixObjectPathIfNecessary(this, value);
			this.context.trackedObjects._autoTrackIfNecessaryWhenHandleObjectResultValue(this, value);
		};
		ClientObject.prototype._handleRetrieveResult=function (value, result) {
			this._handleIdResult(value);
		};
		ClientObject.prototype._recursivelySet=function (input, options, scalarWriteablePropertyNames, objectPropertyNames, notAllowedToBeSetPropertyNames) {
			var isClientObject=(input instanceof ClientObject);
			var originalInput=input;
			if (isClientObject) {
				if (Object.getPrototypeOf(this)===Object.getPrototypeOf(input)) {
					input=JSON.parse(JSON.stringify(input));
				}
				else {
					throw OfficeExtension._Internal.RuntimeError._createInvalidArgError({
						argumentName: 'properties',
						errorLocation: this._className+".set"
					});
				}
			}
			try {
				var prop;
				for (var i=0; i < scalarWriteablePropertyNames.length; i++) {
					prop=scalarWriteablePropertyNames[i];
					if (input.hasOwnProperty(prop)) {
						if (typeof input[prop] !=="undefined") {
							this[prop]=input[prop];
						}
					}
				}
				for (var i=0; i < objectPropertyNames.length; i++) {
					prop=objectPropertyNames[i];
					if (input.hasOwnProperty(prop)) {
						if (typeof input[prop] !=="undefined") {
							var dataToPassToSet=isClientObject ? originalInput[prop] : input[prop];
							this[prop].set(dataToPassToSet, options);
						}
					}
				}
				var throwOnReadOnly=!isClientObject;
				if (options && !OfficeExtension.Utility.isNullOrUndefined(throwOnReadOnly)) {
					throwOnReadOnly=options.throwOnReadOnly;
				}
				for (var i=0; i < notAllowedToBeSetPropertyNames.length; i++) {
					prop=notAllowedToBeSetPropertyNames[i];
					if (input.hasOwnProperty(prop)) {
						if (typeof input[prop] !=="undefined" && throwOnReadOnly) {
							throw new OfficeExtension._Internal.RuntimeError({
								code: OfficeExtension.ErrorCodes.invalidArgument,
								message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.cannotApplyPropertyThroughSetMethod, prop),
								debugInfo: {
									errorLocation: prop
								}
							});
						}
					}
				}
				for (prop in input) {
					if (scalarWriteablePropertyNames.indexOf(prop) < 0 && objectPropertyNames.indexOf(prop) < 0) {
						var propertyDescriptor=Object.getOwnPropertyDescriptor(Object.getPrototypeOf(this), prop);
						if (!propertyDescriptor) {
							throw new OfficeExtension._Internal.RuntimeError({
								code: OfficeExtension.ErrorCodes.invalidArgument,
								message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.propertyDoesNotExist, prop),
								debugInfo: {
									errorLocation: prop
								}
							});
						}
						if (throwOnReadOnly && !propertyDescriptor.set) {
							throw new OfficeExtension._Internal.RuntimeError({
								code: OfficeExtension.ErrorCodes.invalidArgument,
								message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.attemptingToSetReadOnlyProperty, prop),
								debugInfo: {
									errorLocation: prop
								}
							});
						}
					}
				}
			}
			catch (innerError) {
				throw new OfficeExtension._Internal.RuntimeError({
					code: OfficeExtension.ErrorCodes.invalidArgument,
					message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.invalidArgument, 'properties'),
					debugInfo: {
						errorLocation: this._className+".set"
					},
					innerError: innerError
				});
			}
		};
		ClientObject.prototype._recursivelyUpdate=function (properties) {
			var shouldPolyfill=OfficeExtension._internalConfig.alwaysPolyfillClientObjectUpdateMethod;
			if (!shouldPolyfill) {
				shouldPolyfill=!OfficeExtension.Utility.isSetSupported("RichApiRuntime", "1.2");
			}
			try {
				var scalarPropNames=this[OfficeExtension.Constants.scalarPropertyNames];
				if (!scalarPropNames) {
					scalarPropNames=[];
				}
				var scalarPropUpdatable=this[OfficeExtension.Constants.scalarPropertyUpdateable];
				if (!scalarPropUpdatable) {
					scalarPropUpdatable=[];
					for (var i=0; i < scalarPropNames.length; i++) {
						scalarPropUpdatable.push(false);
					}
				}
				var navigationPropNames=this[OfficeExtension.Constants.navigationPropertyNames];
				if (!navigationPropNames) {
					navigationPropNames=[];
				}
				var scalarProps={};
				var navigationProps={};
				var scalarPropCount=0;
				for (var propName in properties) {
					var index=scalarPropNames.indexOf(propName);
					if (index >=0) {
						if (!scalarPropUpdatable[index]) {
							throw new OfficeExtension._Internal.RuntimeError({
								code: OfficeExtension.ErrorCodes.invalidArgument,
								message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.attemptingToSetReadOnlyProperty, propName),
								debugInfo: {
									errorLocation: propName
								}
							});
						}
						scalarProps[propName]=properties[propName];
++scalarPropCount;
					}
					else if (navigationPropNames.indexOf(propName) >=0) {
						navigationProps[propName]=properties[propName];
					}
					else {
						throw new OfficeExtension._Internal.RuntimeError({
							code: OfficeExtension.ErrorCodes.invalidArgument,
							message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.propertyDoesNotExist, propName),
							debugInfo: {
								errorLocation: propName
							}
						});
					}
				}
				if (scalarPropCount > 0) {
					if (shouldPolyfill) {
						for (var i=0; i < scalarPropNames.length; i++) {
							var propName=scalarPropNames[i];
							var propValue=scalarProps[propName];
							if (!OfficeExtension.Utility.isUndefined(propValue)) {
								OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, propName, propValue);
							}
						}
					}
					else {
						OfficeExtension.ActionFactory.createUpdateAction(this.context, this, scalarProps);
					}
				}
				for (var propName in navigationProps) {
					var navigationPropProxy=this[propName];
					var navigationPropValue=navigationProps[propName];
					navigationPropProxy._recursivelyUpdate(navigationPropValue);
				}
			}
			catch (innerError) {
				throw new OfficeExtension._Internal.RuntimeError({
					code: OfficeExtension.ErrorCodes.invalidArgument,
					message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.invalidArgument, 'properties'),
					debugInfo: {
						errorLocation: this._className+".update"
					},
					innerError: innerError
				});
			}
		};
		return ClientObject;
	}());
	OfficeExtension.ClientObject=ClientObject;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ClientRequest=(function () {
		function ClientRequest(context) {
			this.m_context=context;
			this.m_actions=[];
			this.m_actionResultHandler={};
			this.m_referencedObjectPaths={};
			this.m_instantiatedObjectPaths={};
			this.m_flags=0;
			this.m_traceInfos={};
			this.m_pendingProcessEventHandlers=[];
			this.m_pendingEventHandlerActions={};
			this.m_responseTraceIds={};
			this.m_responseTraceMessages=[];
			this.m_preSyncPromises=[];
		}
		Object.defineProperty(ClientRequest.prototype, "flags", {
			get: function () {
				return this.m_flags;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequest.prototype, "traceInfos", {
			get: function () {
				return this.m_traceInfos;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequest.prototype, "_responseTraceMessages", {
			get: function () {
				return this.m_responseTraceMessages;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequest.prototype, "_responseTraceIds", {
			get: function () {
				return this.m_responseTraceIds;
			},
			enumerable: true,
			configurable: true
		});
		ClientRequest.prototype._setResponseTraceIds=function (value) {
			if (value) {
				for (var i=0; i < value.length; i++) {
					var traceId=value[i];
					this.m_responseTraceIds[traceId]=traceId;
					var message=this.m_traceInfos[traceId];
					if (!OfficeExtension.Utility.isNullOrUndefined(message)) {
						this.m_responseTraceMessages.push(message);
					}
				}
			}
		};
		ClientRequest.prototype.addAction=function (action) {
			if (this.m_context.batchMode===1) {
				var isSafeAction=false;
				if (action.actionInfo.ActionType===1 &&
					action.referencedObjectPath.objectPathInfo.ObjectPathType===4) {
					isSafeAction=true;
				}
				if (!isSafeAction) {
					this.m_context.ensureInProgressBatchIfBatchMode();
				}
			}
			if (action.isWriteOperation) {
				this.m_flags=this.m_flags | 1;
			}
			if (action.isRestrictedResourceAccess) {
				this.m_flags=this.m_flags | 2;
			}
			this.m_actions.push(action);
			if (action.actionInfo.ActionType==1) {
				this.m_instantiatedObjectPaths[action.actionInfo.ObjectPathId]=action;
			}
		};
		Object.defineProperty(ClientRequest.prototype, "hasActions", {
			get: function () {
				return this.m_actions.length > 0;
			},
			enumerable: true,
			configurable: true
		});
		ClientRequest.prototype._getLastAction=function () {
			return this.m_actions[this.m_actions.length - 1];
		};
		ClientRequest.prototype.addTrace=function (actionId, message) {
			this.m_traceInfos[actionId]=message;
		};
		ClientRequest.prototype.ensureInstantiateObjectPath=function (objectPath) {
			if (objectPath) {
				if (this.m_instantiatedObjectPaths[objectPath.objectPathInfo.Id]) {
					return;
				}
				this.ensureInstantiateObjectPath(objectPath.parentObjectPath);
				this.ensureInstantiateObjectPaths(objectPath.argumentObjectPaths);
				if (!this.m_instantiatedObjectPaths[objectPath.objectPathInfo.Id]) {
					var actionInfo={
						Id: this.m_context._nextId(),
						ActionType: 1,
						Name: "",
						ObjectPathId: objectPath.objectPathInfo.Id
					};
					var instantiateAction=new OfficeExtension.Action(actionInfo, false, false);
					instantiateAction.referencedObjectPath=objectPath;
					this.addReferencedObjectPath(objectPath);
					this.addAction(instantiateAction);
				}
			}
		};
		ClientRequest.prototype.ensureInstantiateObjectPaths=function (objectPaths) {
			if (objectPaths) {
				for (var i=0; i < objectPaths.length; i++) {
					this.ensureInstantiateObjectPath(objectPaths[i]);
				}
			}
		};
		ClientRequest.prototype.addReferencedObjectPath=function (objectPath) {
			if (this.m_referencedObjectPaths[objectPath.objectPathInfo.Id]) {
				return;
			}
			if (!objectPath.isValid) {
				throw new OfficeExtension._Internal.RuntimeError({
					code: OfficeExtension.ErrorCodes.invalidObjectPath,
					message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.invalidObjectPath, OfficeExtension.Utility.getObjectPathExpression(objectPath)),
					debugInfo: {
						errorLocation: OfficeExtension.Utility.getObjectPathExpression(objectPath)
					}
				});
			}
			while (objectPath) {
				if (objectPath.isWriteOperation) {
					this.m_flags=this.m_flags | 1;
				}
				if (objectPath.isRestrictedResourceAccess) {
					this.m_flags=this.m_flags | 2;
				}
				this.m_referencedObjectPaths[objectPath.objectPathInfo.Id]=objectPath;
				if (objectPath.objectPathInfo.ObjectPathType==3) {
					this.addReferencedObjectPaths(objectPath.argumentObjectPaths);
				}
				objectPath=objectPath.parentObjectPath;
			}
		};
		ClientRequest.prototype.addReferencedObjectPaths=function (objectPaths) {
			if (objectPaths) {
				for (var i=0; i < objectPaths.length; i++) {
					this.addReferencedObjectPath(objectPaths[i]);
				}
			}
		};
		ClientRequest.prototype.addActionResultHandler=function (action, resultHandler) {
			this.m_actionResultHandler[action.actionInfo.Id]=resultHandler;
		};
		ClientRequest.prototype.buildRequestMessageBody=function () {
			if (OfficeExtension._internalConfig.enableEarlyDispose) {
				ClientRequest._calculateLastUsedObjectPathIds(this.m_actions);
			}
			var objectPaths={};
			for (var i in this.m_referencedObjectPaths) {
				objectPaths[i]=this.m_referencedObjectPaths[i].objectPathInfo;
			}
			var actions=[];
			for (var index=0; index < this.m_actions.length; index++) {
				actions.push(this.m_actions[index].actionInfo);
			}
			var ret={
				AutoKeepReference: this.m_context._autoCleanup,
				Actions: actions,
				ObjectPaths: objectPaths
			};
			return ret;
		};
		ClientRequest.prototype.processResponse=function (actionResults) {
			if (actionResults) {
				for (var i=0; i < actionResults.length; i++) {
					var actionResult=actionResults[i];
					var handler=this.m_actionResultHandler[actionResult.ActionId];
					if (handler) {
						handler._handleResult(actionResult.Value);
					}
				}
			}
		};
		ClientRequest.prototype.invalidatePendingInvalidObjectPaths=function () {
			for (var i in this.m_referencedObjectPaths) {
				if (this.m_referencedObjectPaths[i].isInvalidAfterRequest) {
					this.m_referencedObjectPaths[i].isValid=false;
				}
			}
		};
		ClientRequest.prototype._addPendingEventHandlerAction=function (eventHandlers, action) {
			if (!this.m_pendingEventHandlerActions[eventHandlers._id]) {
				this.m_pendingEventHandlerActions[eventHandlers._id]=[];
				this.m_pendingProcessEventHandlers.push(eventHandlers);
			}
			this.m_pendingEventHandlerActions[eventHandlers._id].push(action);
		};
		Object.defineProperty(ClientRequest.prototype, "_pendingProcessEventHandlers", {
			get: function () {
				return this.m_pendingProcessEventHandlers;
			},
			enumerable: true,
			configurable: true
		});
		ClientRequest.prototype._getPendingEventHandlerActions=function (eventHandlers) {
			return this.m_pendingEventHandlerActions[eventHandlers._id];
		};
		ClientRequest.prototype._addPreSyncPromise=function (value) {
			this.m_preSyncPromises.push(value);
		};
		Object.defineProperty(ClientRequest.prototype, "_preSyncPromises", {
			get: function () {
				return this.m_preSyncPromises;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequest.prototype, "_actions", {
			get: function () {
				return this.m_actions;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequest.prototype, "_objectPaths", {
			get: function () {
				return this.m_referencedObjectPaths;
			},
			enumerable: true,
			configurable: true
		});
		ClientRequest._updateLastUsedActionIdOfObjectPathId=function (lastUsedActionIdOfObjectPathId, objectPath, actionId) {
			while (objectPath) {
				if (lastUsedActionIdOfObjectPathId[objectPath.objectPathInfo.Id]) {
					return;
				}
				lastUsedActionIdOfObjectPathId[objectPath.objectPathInfo.Id]=actionId;
				var argumentObjectPaths=objectPath.argumentObjectPaths;
				if (argumentObjectPaths) {
					var argumentObjectPathsLength=argumentObjectPaths.length;
					for (var i=0; i < argumentObjectPathsLength; i++) {
						ClientRequest._updateLastUsedActionIdOfObjectPathId(lastUsedActionIdOfObjectPathId, argumentObjectPaths[i], actionId);
					}
				}
				objectPath=objectPath.parentObjectPath;
			}
		};
		ClientRequest._calculateLastUsedObjectPathIds=function (actions) {
			var lastUsedActionIdOfObjectPathId={};
			var actionsLength=actions.length;
			for (var index=actionsLength - 1; index >=0; --index) {
				var action=actions[index];
				var actionId=action.actionInfo.Id;
				if (action.referencedObjectPath) {
					ClientRequest._updateLastUsedActionIdOfObjectPathId(lastUsedActionIdOfObjectPathId, action.referencedObjectPath, actionId);
				}
				var referencedObjectPaths=action.referencedArgumentObjectPaths;
				if (referencedObjectPaths) {
					var referencedObjectPathsLength=referencedObjectPaths.length;
					for (var refIndex=0; refIndex < referencedObjectPathsLength; refIndex++) {
						ClientRequest._updateLastUsedActionIdOfObjectPathId(lastUsedActionIdOfObjectPathId, referencedObjectPaths[refIndex], actionId);
					}
				}
			}
			var lastUsedObjectPathIdsOfAction={};
			for (var key in lastUsedActionIdOfObjectPathId) {
				var actionId=lastUsedActionIdOfObjectPathId[key];
				var objectPathIds=lastUsedObjectPathIdsOfAction[actionId];
				if (!objectPathIds) {
					objectPathIds=[];
					lastUsedObjectPathIdsOfAction[actionId]=objectPathIds;
				}
				objectPathIds.push(parseInt(key));
			}
			for (var index=0; index < actionsLength; index++) {
				var action=actions[index];
				var lastUsedObjectPathIds=lastUsedObjectPathIdsOfAction[action.actionInfo.Id];
				if (lastUsedObjectPathIds && lastUsedObjectPathIds.length > 0) {
					action.actionInfo.L=lastUsedObjectPathIds;
				}
				else if (action.actionInfo.L) {
					delete action.actionInfo.L;
				}
			}
		};
		return ClientRequest;
	}());
	OfficeExtension.ClientRequest=ClientRequest;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	OfficeExtension._internalConfig={
		showDisposeInfoInDebugInfo: false,
		enableEarlyDispose: true,
		alwaysPolyfillClientObjectUpdateMethod: false,
		alwaysPolyfillClientObjectRetrieveMethod: false
	};
	var SessionBase=(function () {
		function SessionBase() {
		}
		SessionBase.prototype._resolveRequestUrlAndHeaderInfo=function () {
			return OfficeExtension.Utility._createPromiseFromResult(null);
		};
		SessionBase.prototype._createRequestExecutorOrNull=function () {
			return null;
		};
		Object.defineProperty(SessionBase.prototype, "eventRegistration", {
			get: function () {
				return OfficeExtension._Internal.officeJsEventRegistration;
			},
			enumerable: true,
			configurable: true
		});
		return SessionBase;
	}());
	OfficeExtension.SessionBase=SessionBase;
	var ClientRequestContext=(function () {
		function ClientRequestContext(url) {
			this.m_customRequestHeaders={};
			this.m_batchMode=0;
			this._onRunFinishedNotifiers=[];
			this.m_nextId=0;
			if (ClientRequestContext._overrideSession) {
				this.m_requestUrlAndHeaderInfoResolver=ClientRequestContext._overrideSession;
			}
			else {
				if (OfficeExtension.Utility.isNullOrUndefined(url) || typeof (url)==="string" && url.length===0) {
					url=ClientRequestContext.defaultRequestUrlAndHeaders;
					if (!url) {
						url={ url: OfficeExtension.Constants.localDocument, headers: {} };
					}
				}
				if (typeof (url)==="string") {
					this.m_requestUrlAndHeaderInfo={ url: url, headers: {} };
				}
				else if (ClientRequestContext.isRequestUrlAndHeaderInfoResolver(url)) {
					this.m_requestUrlAndHeaderInfoResolver=url;
				}
				else if (ClientRequestContext.isRequestUrlAndHeaderInfo(url)) {
					var requestInfo=url;
					this.m_requestUrlAndHeaderInfo={ url: requestInfo.url, headers: {} };
					OfficeExtension.Utility._copyHeaders(requestInfo.headers, this.m_requestUrlAndHeaderInfo.headers);
				}
				else {
					throw OfficeExtension._Internal.RuntimeError._createInvalidArgError("url");
				}
			}
			if (this.m_requestUrlAndHeaderInfoResolver instanceof SessionBase) {
				this.m_session=this.m_requestUrlAndHeaderInfoResolver;
			}
			this._processingResult=false;
			this._customData=OfficeExtension.Constants.iterativeExecutor;
			this.sync=this.sync.bind(this);
		}
		Object.defineProperty(ClientRequestContext.prototype, "session", {
			get: function () {
				return this.m_session;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequestContext.prototype, "eventRegistration", {
			get: function () {
				if (this.m_session) {
					return this.m_session.eventRegistration;
				}
				return OfficeExtension._Internal.officeJsEventRegistration;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequestContext.prototype, "_url", {
			get: function () {
				if (this.m_requestUrlAndHeaderInfo) {
					return this.m_requestUrlAndHeaderInfo.url;
				}
				return null;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequestContext.prototype, "_pendingRequest", {
			get: function () {
				if (this.m_pendingRequest==null) {
					this.m_pendingRequest=new OfficeExtension.ClientRequest(this);
				}
				return this.m_pendingRequest;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequestContext.prototype, "debugInfo", {
			get: function () {
				var prettyPrinter=new OfficeExtension.RequestPrettyPrinter(this._rootObjectPropertyName, this._pendingRequest._objectPaths, this._pendingRequest._actions, OfficeExtension._internalConfig.showDisposeInfoInDebugInfo);
				var statements=prettyPrinter.process();
				return { pendingStatements: statements };
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequestContext.prototype, "trackedObjects", {
			get: function () {
				if (!this.m_trackedObjects) {
					this.m_trackedObjects=new OfficeExtension.TrackedObjects(this);
				}
				return this.m_trackedObjects;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequestContext.prototype, "requestHeaders", {
			get: function () {
				return this.m_customRequestHeaders;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ClientRequestContext.prototype, "batchMode", {
			get: function () {
				return this.m_batchMode;
			},
			enumerable: true,
			configurable: true
		});
		ClientRequestContext.prototype.ensureInProgressBatchIfBatchMode=function () {
			if (this.m_batchMode===1 && !this.m_explicitBatchInProgress) {
				throw OfficeExtension.Utility.createRuntimeError(OfficeExtension.ErrorCodes.generalException, OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.notInsideBatch), null);
			}
		};
		ClientRequestContext.prototype.load=function (clientObj, option) {
			OfficeExtension.Utility.validateContext(this, clientObj);
			var queryOption=ClientRequestContext._parseQueryOption(option);
			var action=OfficeExtension.ActionFactory.createQueryAction(this, clientObj, queryOption);
			this._pendingRequest.addActionResultHandler(action, clientObj);
		};
		ClientRequestContext.isLoadOption=function (loadOption) {
			if (!OfficeExtension.Utility.isUndefined(loadOption.select) && (typeof (loadOption.select)==="string" || Array.isArray(loadOption.select)))
				return true;
			if (!OfficeExtension.Utility.isUndefined(loadOption.expand) && (typeof (loadOption.expand)==="string" || Array.isArray(loadOption.expand)))
				return true;
			if (!OfficeExtension.Utility.isUndefined(loadOption.top) && typeof (loadOption.top)==="number")
				return true;
			if (!OfficeExtension.Utility.isUndefined(loadOption.skip) && typeof (loadOption.skip)==="number")
				return true;
			for (var i in loadOption) {
				return false;
			}
			return true;
		};
		ClientRequestContext.parseStrictLoadOption=function (option) {
			var ret={ Select: [] };
			ClientRequestContext.parseStrictLoadOptionHelper(ret, "", "option", option);
			return ret;
		};
		ClientRequestContext.combineQueryPath=function (pathPrefix, key, separator) {
			if (pathPrefix.length===0) {
				return key;
			}
			else {
				return pathPrefix+separator+key;
			}
		};
		ClientRequestContext.parseStrictLoadOptionHelper=function (queryInfo, pathPrefix, argPrefix, option) {
			for (var key in option) {
				var value=option[key];
				if (key==="$all") {
					if (typeof (value) !=="boolean") {
						throw OfficeExtension._Internal.RuntimeError._createInvalidArgError(ClientRequestContext.combineQueryPath(argPrefix, key, "."));
					}
					if (value) {
						queryInfo.Select.push(ClientRequestContext.combineQueryPath(pathPrefix, "*", "/"));
					}
				}
				else if (key==="$top") {
					if (typeof (value) !=="number" || pathPrefix.length > 0) {
						throw OfficeExtension._Internal.RuntimeError._createInvalidArgError(ClientRequestContext.combineQueryPath(argPrefix, key, "."));
					}
					queryInfo.Top=value;
				}
				else if (key==="$skip") {
					if (typeof (value) !=="number" || pathPrefix.length > 0) {
						throw OfficeExtension._Internal.RuntimeError._createInvalidArgError(ClientRequestContext.combineQueryPath(argPrefix, key, "."));
					}
					queryInfo.Skip=value;
				}
				else {
					if (typeof (value)==="boolean") {
						if (value) {
							queryInfo.Select.push(ClientRequestContext.combineQueryPath(pathPrefix, key, "/"));
						}
					}
					else if (typeof (value)==="object") {
						ClientRequestContext.parseStrictLoadOptionHelper(queryInfo, ClientRequestContext.combineQueryPath(pathPrefix, key, "/"), ClientRequestContext.combineQueryPath(argPrefix, key, "."), value);
					}
					else {
						throw OfficeExtension._Internal.RuntimeError._createInvalidArgError(ClientRequestContext.combineQueryPath(argPrefix, key, "."));
					}
				}
			}
		};
		ClientRequestContext._parseQueryOption=function (option) {
			var queryOption={};
			if (typeof (option)=="string") {
				var select=option;
				queryOption.Select=OfficeExtension.Utility._parseSelectExpand(select);
			}
			else if (Array.isArray(option)) {
				queryOption.Select=option;
			}
			else if (typeof (option)==="object") {
				var loadOption=option;
				if (ClientRequestContext.isLoadOption(loadOption)) {
					if (typeof (loadOption.select)=="string") {
						queryOption.Select=OfficeExtension.Utility._parseSelectExpand(loadOption.select);
					}
					else if (Array.isArray(loadOption.select)) {
						queryOption.Select=loadOption.select;
					}
					else if (!OfficeExtension.Utility.isNullOrUndefined(loadOption.select)) {
						throw OfficeExtension._Internal.RuntimeError._createInvalidArgError("option.select");
					}
					if (typeof (loadOption.expand)=="string") {
						queryOption.Expand=OfficeExtension.Utility._parseSelectExpand(loadOption.expand);
					}
					else if (Array.isArray(loadOption.expand)) {
						queryOption.Expand=loadOption.expand;
					}
					else if (!OfficeExtension.Utility.isNullOrUndefined(loadOption.expand)) {
						throw OfficeExtension._Internal.RuntimeError._createInvalidArgError("option.expand");
					}
					if (typeof (loadOption.top)==="number") {
						queryOption.Top=loadOption.top;
					}
					else if (!OfficeExtension.Utility.isNullOrUndefined(loadOption.top)) {
						throw OfficeExtension._Internal.RuntimeError._createInvalidArgError("option.top");
					}
					if (typeof (loadOption.skip)==="number") {
						queryOption.Skip=loadOption.skip;
					}
					else if (!OfficeExtension.Utility.isNullOrUndefined(loadOption.skip)) {
						throw OfficeExtension._Internal.RuntimeError._createInvalidArgError("option.skip");
					}
				}
				else {
					queryOption=ClientRequestContext.parseStrictLoadOption(option);
				}
			}
			else if (!OfficeExtension.Utility.isNullOrUndefined(option)) {
				throw OfficeExtension._Internal.RuntimeError._createInvalidArgError("option");
			}
			return queryOption;
		};
		ClientRequestContext.prototype.loadRecursive=function (clientObj, options, maxDepth) {
			if (!OfficeExtension.Utility.isPlainJsonObject(options)) {
				throw OfficeExtension._Internal.RuntimeError._createInvalidArgError("options");
			}
			var quries={};
			for (var key in options) {
				quries[key]=ClientRequestContext._parseQueryOption(options[key]);
			}
			var action=OfficeExtension.ActionFactory.createRecursiveQueryAction(this, clientObj, { Queries: quries, MaxDepth: maxDepth });
			this._pendingRequest.addActionResultHandler(action, clientObj);
		};
		ClientRequestContext.prototype.trace=function (message) {
			OfficeExtension.ActionFactory.createTraceAction(this, message, true);
		};
		ClientRequestContext.prototype._processOfficeJsErrorResponse=function (officeJsErrorCode, response) {
		};
		ClientRequestContext.prototype.ensureRequestUrlAndHeaderInfo=function () {
			var _this=this;
			return OfficeExtension.Utility._createPromiseFromResult(null)
				.then(function () {
				if (!_this.m_requestUrlAndHeaderInfo) {
					return _this.m_requestUrlAndHeaderInfoResolver._resolveRequestUrlAndHeaderInfo()
						.then(function (value) {
						_this.m_requestUrlAndHeaderInfo=value;
						if (!_this.m_requestUrlAndHeaderInfo) {
							_this.m_requestUrlAndHeaderInfo={ url: OfficeExtension.Constants.localDocument, headers: {} };
						}
						if (OfficeExtension.Utility.isNullOrEmptyString(_this.m_requestUrlAndHeaderInfo.url)) {
							_this.m_requestUrlAndHeaderInfo.url=OfficeExtension.Constants.localDocument;
						}
						if (!_this.m_requestUrlAndHeaderInfo.headers) {
							_this.m_requestUrlAndHeaderInfo.headers={};
						}
						if (typeof (_this.m_requestUrlAndHeaderInfoResolver._createRequestExecutorOrNull)==="function") {
							var executor=_this.m_requestUrlAndHeaderInfoResolver._createRequestExecutorOrNull();
							if (executor) {
								_this._requestExecutor=executor;
							}
						}
					});
				}
			});
		};
		ClientRequestContext.prototype.syncPrivateMain=function () {
			var _this=this;
			return this.ensureRequestUrlAndHeaderInfo()
				.then(function () {
				var req=_this._pendingRequest;
				_this.m_pendingRequest=null;
				return _this.processPreSyncPromises(req)
					.then(function () { return _this.syncPrivate(req); });
			});
		};
		ClientRequestContext.prototype.syncPrivate=function (req) {
			var _this=this;
			if (!req.hasActions) {
				return this.processPendingEventHandlers(req);
			}
			var msgBody=req.buildRequestMessageBody();
			var requestFlags=req.flags;
			if (!this._requestExecutor) {
				if (OfficeExtension.Utility._isLocalDocumentUrl(this.m_requestUrlAndHeaderInfo.url)) {
					this._requestExecutor=new OfficeExtension.OfficeJsRequestExecutor(this);
				}
				else {
					this._requestExecutor=new OfficeExtension.HttpRequestExecutor();
				}
			}
			var requestExecutor=this._requestExecutor;
			var headers={};
			OfficeExtension.Utility._copyHeaders(this.m_requestUrlAndHeaderInfo.headers, headers);
			OfficeExtension.Utility._copyHeaders(this.m_customRequestHeaders, headers);
			var requestExecutorRequestMessage={
				Url: this.m_requestUrlAndHeaderInfo.url,
				Headers: headers,
				Body: msgBody
			};
			req.invalidatePendingInvalidObjectPaths();
			var errorFromResponse=null;
			var errorFromProcessEventHandlers=null;
			this._lastSyncStart=performance.now();
			return requestExecutor.executeAsync(this._customData, requestFlags, requestExecutorRequestMessage)
				.then(function (response) {
				_this._lastSyncEnd=performance.now();
				errorFromResponse=_this.processRequestExecutorResponseMessage(req, response);
				return _this.processPendingEventHandlers(req)
					.catch(function (ex) {
					OfficeExtension.Utility.log("Error in processPendingEventHandlers");
					OfficeExtension.Utility.log(JSON.stringify(ex));
					errorFromProcessEventHandlers=ex;
				});
			})
				.then(function () {
				if (errorFromResponse) {
					OfficeExtension.Utility.log("Throw error from response: "+JSON.stringify(errorFromResponse));
					throw errorFromResponse;
				}
				if (errorFromProcessEventHandlers) {
					OfficeExtension.Utility.log("Throw error from ProcessEventHandler: "+JSON.stringify(errorFromProcessEventHandlers));
					var transformedError=null;
					if (errorFromProcessEventHandlers instanceof OfficeExtension._Internal.RuntimeError) {
						transformedError=errorFromProcessEventHandlers;
						transformedError.traceMessages=req._responseTraceMessages;
					}
					else {
						var message=null;
						if (typeof (errorFromProcessEventHandlers)==="string") {
							message=errorFromProcessEventHandlers;
						}
						else {
							message=errorFromProcessEventHandlers.message;
						}
						if (OfficeExtension.Utility.isNullOrEmptyString(message)) {
							message=OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.cannotRegisterEvent);
						}
						transformedError=new OfficeExtension._Internal.RuntimeError({
							code: OfficeExtension.ErrorCodes.cannotRegisterEvent,
							message: message,
							traceMessages: req._responseTraceMessages
						});
					}
					throw transformedError;
				}
			});
		};
		ClientRequestContext.prototype.processRequestExecutorResponseMessage=function (req, response) {
			if (response.Body && response.Body.TraceIds) {
				req._setResponseTraceIds(response.Body.TraceIds);
			}
			var traceMessages=req._responseTraceMessages;
			var errorStatementInfo=null;
			if (response.Body) {
				if (response.Body.Error &&
					response.Body.Error.ActionIndex >=0) {
					var prettyPrinter=new OfficeExtension.RequestPrettyPrinter(this._rootObjectPropertyName, req._objectPaths, req._actions);
					errorStatementInfo=prettyPrinter.processForDebugStatementInfo(response.Body.Error.ActionIndex);
				}
				var actionResults=null;
				if (response.Body.Results) {
					actionResults=response.Body.Results;
				}
				else if (response.Body.ProcessedResults && response.Body.ProcessedResults.Results) {
					actionResults=response.Body.ProcessedResults.Results;
				}
				if (actionResults) {
					this._processingResult=true;
					try {
						req.processResponse(actionResults);
					}
					finally {
						this._processingResult=false;
					}
				}
			}
			if (!OfficeExtension.Utility.isNullOrEmptyString(response.ErrorCode)) {
				return new OfficeExtension._Internal.RuntimeError({
					code: response.ErrorCode,
					message: response.ErrorMessage,
					traceMessages: traceMessages
				});
			}
			else if (response.Body && response.Body.Error) {
				var debugInfo={
					errorLocation: response.Body.Error.Location
				};
				if (errorStatementInfo) {
					debugInfo.statement=errorStatementInfo.statement;
					debugInfo.surroundingStatements=errorStatementInfo.surroundingStatements;
				}
				return new OfficeExtension._Internal.RuntimeError({
					code: response.Body.Error.Code,
					message: response.Body.Error.Message,
					traceMessages: traceMessages,
					debugInfo: debugInfo
				});
			}
			return null;
		};
		ClientRequestContext.prototype.processPendingEventHandlers=function (req) {
			var ret=OfficeExtension.Utility._createPromiseFromResult(null);
			for (var i=0; i < req._pendingProcessEventHandlers.length; i++) {
				var eventHandlers=req._pendingProcessEventHandlers[i];
				ret=ret.then(this.createProcessOneEventHandlersFunc(eventHandlers, req));
			}
			return ret;
		};
		ClientRequestContext.prototype.createProcessOneEventHandlersFunc=function (eventHandlers, req) {
			return function () { return eventHandlers._processRegistration(req); };
		};
		ClientRequestContext.prototype.processPreSyncPromises=function (req) {
			var ret=OfficeExtension.Utility._createPromiseFromResult(null);
			for (var i=0; i < req._preSyncPromises.length; i++) {
				var p=req._preSyncPromises[i];
				ret=ret.then(this.createProcessOneProSyncFunc(p));
			}
			return ret;
		};
		ClientRequestContext.prototype.createProcessOneProSyncFunc=function (p) {
			return function () { return p; };
		};
		ClientRequestContext.prototype.sync=function (passThroughValue) {
			return this.syncPrivateMain().then(function () { return passThroughValue; });
		};
		ClientRequestContext.prototype.batch=function (batchBody) {
			var _this=this;
			if (this.m_batchMode !==1) {
				return OfficeExtension._Internal.OfficePromise.reject(OfficeExtension.Utility.createRuntimeError(OfficeExtension.ErrorCodes.generalException, null, null));
			}
			if (this.m_explicitBatchInProgress) {
				return OfficeExtension._Internal.OfficePromise.reject(OfficeExtension.Utility.createRuntimeError(OfficeExtension.ErrorCodes.generalException, OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.pendingBatchInProgress), null));
			}
			if (OfficeExtension.Utility.isNullOrUndefined(batchBody)) {
				return OfficeExtension.Utility._createPromiseFromResult(null);
			}
			this.m_explicitBatchInProgress=true;
			var previousRequest=this.m_pendingRequest;
			this.m_pendingRequest=new OfficeExtension.ClientRequest(this);
			var batchBodyResult;
			try {
				batchBodyResult=batchBody(this._rootObject, this);
			}
			catch (ex) {
				this.m_explicitBatchInProgress=false;
				this.m_pendingRequest=previousRequest;
				return OfficeExtension._Internal.OfficePromise.reject(ex);
			}
			var request;
			var batchBodyResultPromise;
			if (typeof (batchBodyResult)==="object" &&
				batchBodyResult &&
				typeof (batchBodyResult.then)==="function") {
				batchBodyResultPromise=OfficeExtension.Utility._createPromiseFromResult(null)
					.then(function () {
					return batchBodyResult;
				})
					.then(function (result) {
					_this.m_explicitBatchInProgress=false;
					request=_this.m_pendingRequest;
					_this.m_pendingRequest=previousRequest;
					return result;
				})
					.catch(function (ex) {
					_this.m_explicitBatchInProgress=false;
					request=_this.m_pendingRequest;
					_this.m_pendingRequest=previousRequest;
					return OfficeExtension._Internal.OfficePromise.reject(ex);
				});
			}
			else {
				this.m_explicitBatchInProgress=false;
				request=this.m_pendingRequest;
				this.m_pendingRequest=previousRequest;
				batchBodyResultPromise=OfficeExtension.Utility._createPromiseFromResult(batchBodyResult);
			}
			return batchBodyResultPromise
				.then(function (result) {
				return _this.ensureRequestUrlAndHeaderInfo()
					.then(function () {
					return _this.syncPrivate(request);
				})
					.then(function () {
					return result;
				});
			});
		};
		ClientRequestContext._run=function (ctxInitializer, runBody, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
			if (numCleanupAttempts===void 0) { numCleanupAttempts=3; }
			if (retryDelay===void 0) { retryDelay=5000; }
			return ClientRequestContext._runCommon("run", null, ctxInitializer, 0, runBody, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure);
		};
		ClientRequestContext.isRequestUrlAndHeaderInfo=function (value) {
			return (typeof (value)==="object" &&
				value !==null &&
				Object.getPrototypeOf(value)===Object.getPrototypeOf({}) &&
				!OfficeExtension.Utility.isNullOrUndefined(value.url));
		};
		ClientRequestContext.isRequestUrlAndHeaderInfoResolver=function (value) {
			return (typeof (value)==="object" &&
				value !==null &&
				typeof (value._resolveRequestUrlAndHeaderInfo)==="function");
		};
		ClientRequestContext._runBatch=function (functionName, receivedRunArgs, ctxInitializer, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
			if (numCleanupAttempts===void 0) { numCleanupAttempts=3; }
			if (retryDelay===void 0) { retryDelay=5000; }
			return ClientRequestContext._runBatchCommon(0, functionName, receivedRunArgs, ctxInitializer, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure);
		};
		ClientRequestContext._runExplicitBatch=function (functionName, receivedRunArgs, ctxInitializer, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
			if (numCleanupAttempts===void 0) { numCleanupAttempts=3; }
			if (retryDelay===void 0) { retryDelay=5000; }
			return ClientRequestContext._runBatchCommon(1, functionName, receivedRunArgs, ctxInitializer, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure);
		};
		ClientRequestContext._runBatchCommon=function (batchMode, functionName, receivedRunArgs, ctxInitializer, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
			if (numCleanupAttempts===void 0) { numCleanupAttempts=3; }
			if (retryDelay===void 0) { retryDelay=5000; }
			var ctxRetriever;
			var batch;
			var requestInfo=null;
			var argOffset=0;
			if (receivedRunArgs.length > 0 &&
				(typeof (receivedRunArgs[0])==="string" ||
					ClientRequestContext.isRequestUrlAndHeaderInfo(receivedRunArgs[0]) ||
					ClientRequestContext.isRequestUrlAndHeaderInfoResolver(receivedRunArgs[0]))) {
				requestInfo=receivedRunArgs[0];
				argOffset=1;
			}
			if (receivedRunArgs.length==argOffset+1) {
				ctxRetriever=ctxInitializer;
				batch=receivedRunArgs[argOffset+0];
			}
			else if (receivedRunArgs.length==argOffset+2) {
				if (OfficeExtension.Utility.isNullOrUndefined(receivedRunArgs[argOffset+0])) {
					ctxRetriever=ctxInitializer;
				}
				else if (receivedRunArgs[argOffset+0] instanceof OfficeExtension.ClientObject) {
					ctxRetriever=function () { return receivedRunArgs[argOffset+0].context; };
				}
				else if (receivedRunArgs[argOffset+0] instanceof ClientRequestContext) {
					ctxRetriever=function () { return receivedRunArgs[argOffset+0]; };
				}
				else if (Array.isArray(receivedRunArgs[argOffset+0])) {
					var array=receivedRunArgs[argOffset+0];
					if (array.length==0) {
						return ClientRequestContext.createErrorPromise(functionName);
					}
					for (var i=0; i < array.length; i++) {
						if (!(array[i] instanceof OfficeExtension.ClientObject)) {
							return ClientRequestContext.createErrorPromise(functionName);
						}
						if (array[i].context !=array[0].context) {
							return ClientRequestContext.createErrorPromise(functionName, OfficeExtension.ResourceStrings.invalidRequestContext);
						}
					}
					ctxRetriever=function () { return array[0].context; };
				}
				else {
					return ClientRequestContext.createErrorPromise(functionName);
				}
				batch=receivedRunArgs[argOffset+1];
			}
			else {
				return ClientRequestContext.createErrorPromise(functionName);
			}
			return ClientRequestContext._runCommon(functionName, requestInfo, ctxRetriever, batchMode, batch, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure);
		};
		ClientRequestContext.createErrorPromise=function (functionName, code) {
			if (code===void 0) { code=OfficeExtension.ResourceStrings.invalidArgument; }
			return OfficeExtension._Internal.OfficePromise.reject(OfficeExtension.Utility.createRuntimeError(code, OfficeExtension.Utility._getResourceString(code), functionName));
		};
		ClientRequestContext._runCommon=function (functionName, requestInfo, ctxRetriever, batchMode, runBody, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
			if (ClientRequestContext._overrideSession) {
				requestInfo=ClientRequestContext._overrideSession;
			}
			var starterPromise=new OfficeExtension._Internal.OfficePromise(function (resolve, reject) { resolve(); });
			var ctx;
			var succeeded=false;
			var resultOrError;
			var previousBatchMode;
			return starterPromise
				.then(function () {
				ctx=ctxRetriever(requestInfo);
				if (ctx._autoCleanup) {
					return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
						ctx._onRunFinishedNotifiers.push(function () {
							ctx._autoCleanup=true;
							resolve();
						});
					});
				}
				else {
					ctx._autoCleanup=true;
				}
			})
				.then(function () {
				if (typeof runBody !=='function') {
					return ClientRequestContext.createErrorPromise(functionName);
				}
				previousBatchMode=ctx.m_batchMode;
				ctx.m_batchMode=batchMode;
				var runBodyResult;
				if (batchMode==1) {
					runBodyResult=runBody(ctx.batch.bind(ctx));
				}
				else {
					runBodyResult=runBody(ctx);
				}
				if (OfficeExtension.Utility.isNullOrUndefined(runBodyResult) || (typeof runBodyResult.then !=='function')) {
					OfficeExtension.Utility.throwError(OfficeExtension.ResourceStrings.runMustReturnPromise);
				}
				return runBodyResult;
			})
				.then(function (runBodyResult) {
				if (batchMode===1) {
					return runBodyResult;
				}
				else {
					return ctx.sync(runBodyResult);
				}
			})
				.then(function (result) {
				succeeded=true;
				resultOrError=result;
			})
				.catch(function (error) {
				resultOrError=error;
			})
				.then(function () {
				var itemsToRemove=ctx.trackedObjects._retrieveAndClearAutoCleanupList();
				ctx._autoCleanup=false;
				ctx.m_batchMode=previousBatchMode;
				for (var key in itemsToRemove) {
					itemsToRemove[key]._objectPath.isValid=false;
				}
				var cleanupCounter=0;
				if (OfficeExtension.Utility._synchronousCleanup || ClientRequestContext.isRequestUrlAndHeaderInfoResolver(requestInfo)) {
					return attemptCleanup();
				}
				else {
					attemptCleanup();
				}
				function attemptCleanup() {
					cleanupCounter++;
					var savedPendingRequest=ctx.m_pendingRequest;
					var savedBatchMode=ctx.m_batchMode;
					var request=new OfficeExtension.ClientRequest(ctx);
					ctx.m_pendingRequest=request;
					ctx.m_batchMode=0;
					try {
						for (var key in itemsToRemove) {
							ctx.trackedObjects.remove(itemsToRemove[key]);
						}
					}
					finally {
						ctx.m_batchMode=savedBatchMode;
						ctx.m_pendingRequest=savedPendingRequest;
					}
					return ctx.syncPrivate(request)
						.then(function () {
						if (onCleanupSuccess) {
							onCleanupSuccess(cleanupCounter);
						}
					})
						.catch(function () {
						if (onCleanupFailure) {
							onCleanupFailure(cleanupCounter);
						}
						if (cleanupCounter < numCleanupAttempts) {
							setTimeout(function () {
								attemptCleanup();
							}, retryDelay);
						}
					});
				}
			})
				.then(function () {
				if (ctx._onRunFinishedNotifiers && ctx._onRunFinishedNotifiers.length > 0) {
					var func=ctx._onRunFinishedNotifiers.shift();
					func();
				}
				if (succeeded) {
					return resultOrError;
				}
				else {
					throw resultOrError;
				}
			});
		};
		ClientRequestContext.prototype._nextId=function () {
			return++this.m_nextId;
		};
		return ClientRequestContext;
	}());
	OfficeExtension.ClientRequestContext=ClientRequestContext;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ClientResult=(function () {
		function ClientResult(type) {
			this.m_type=type;
		}
		Object.defineProperty(ClientResult.prototype, "value", {
			get: function () {
				if (!this.m_isLoaded) {
					throw new OfficeExtension._Internal.RuntimeError({
						code: OfficeExtension.ErrorCodes.valueNotLoaded,
						message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.valueNotLoaded),
						debugInfo: {
							errorLocation: "clientResult.value"
						}
					});
				}
				return this.m_value;
			},
			enumerable: true,
			configurable: true
		});
		ClientResult.prototype._handleResult=function (value) {
			this.m_isLoaded=true;
			if (typeof (value)==="object" && value && value._IsNull) {
				return;
			}
			if (this.m_type===1) {
				this.m_value=OfficeExtension.Utility.adjustToDateTime(value);
			}
			else {
				this.m_value=value;
			}
		};
		return ClientResult;
	}());
	OfficeExtension.ClientResult=ClientResult;
	var RetrieveResultImpl=(function () {
		function RetrieveResultImpl(m_proxy, m_shouldPolyfill) {
			this.m_proxy=m_proxy;
			this.m_shouldPolyfill=m_shouldPolyfill;
			var scalarPropertyNames=m_proxy[OfficeExtension.Constants.scalarPropertyNames];
			var navigationPropertyNames=m_proxy[OfficeExtension.Constants.navigationPropertyNames];
			var typeName=m_proxy[OfficeExtension.Constants.className];
			var isCollection=m_proxy[OfficeExtension.Constants.isCollection];
			if (scalarPropertyNames) {
				for (var i=0; i < scalarPropertyNames.length; i++) {
					OfficeExtension.Utility.definePropertyThrowUnloadedException(this, typeName, scalarPropertyNames[i]);
				}
			}
			if (navigationPropertyNames) {
				for (var i=0; i < navigationPropertyNames.length; i++) {
					OfficeExtension.Utility.definePropertyThrowUnloadedException(this, typeName, navigationPropertyNames[i]);
				}
			}
			if (isCollection) {
				OfficeExtension.Utility.definePropertyThrowUnloadedException(this, typeName, OfficeExtension.Constants.itemsLowerCase);
			}
		}
		Object.defineProperty(RetrieveResultImpl.prototype, "$proxy", {
			get: function () {
				return this.m_proxy;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RetrieveResultImpl.prototype, "$isNullObject", {
			get: function () {
				if (!this.m_isLoaded) {
					throw new OfficeExtension._Internal.RuntimeError({
						code: OfficeExtension.ErrorCodes.valueNotLoaded,
						message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.valueNotLoaded),
						debugInfo: {
							errorLocation: "retrieveResult.$isNullObject"
						}
					});
				}
				return this.m_isNullObject;
			},
			enumerable: true,
			configurable: true
		});
		RetrieveResultImpl.prototype.toJSON=function () {
			if (!this.m_isLoaded) {
				return undefined;
			}
			if (this.m_isNullObject) {
				return null;
			}
			if (OfficeExtension.Utility.isUndefined(this.m_json)) {
				this.m_json=this.purifyJson(this.m_value);
			}
			return this.m_json;
		};
		RetrieveResultImpl.prototype.toString=function () {
			return JSON.stringify(this.toJSON());
		};
		RetrieveResultImpl.prototype._handleResult=function (value) {
			this.m_isLoaded=true;
			if (value===null || typeof (value)==="object" && value && value._IsNull) {
				this.m_isNullObject=true;
				value=null;
			}
			else {
				this.m_isNullObject=false;
			}
			if (this.m_shouldPolyfill) {
				value=this.changePropertyNameToCamelLowerCase(value);
			}
			this.m_value=value;
			this.m_proxy._handleRetrieveResult(value, this);
		};
		RetrieveResultImpl.prototype.changePropertyNameToCamelLowerCase=function (value) {
			var charCodeUnderscore=95;
			if (Array.isArray(value)) {
				var ret=[];
				for (var i=0; i < value.length; i++) {
					ret.push(this.changePropertyNameToCamelLowerCase(value[i]));
				}
				return ret;
			}
			else if (typeof (value)==="object" && value !==null) {
				var ret={};
				for (var key in value) {
					var propValue=value[key];
					if (key===OfficeExtension.Constants.items) {
						ret={};
						ret[OfficeExtension.Constants.itemsLowerCase]=this.changePropertyNameToCamelLowerCase(propValue);
						break;
					}
					else {
						var propName=OfficeExtension.Utility._toCamelLowerCase(key);
						ret[propName]=this.changePropertyNameToCamelLowerCase(propValue);
					}
				}
				return ret;
			}
			else {
				return value;
			}
		};
		RetrieveResultImpl.prototype.purifyJson=function (value) {
			var charCodeUnderscore=95;
			if (Array.isArray(value)) {
				var ret=[];
				for (var i=0; i < value.length; i++) {
					ret.push(this.purifyJson(value[i]));
				}
				return ret;
			}
			else if (typeof (value)==="object" && value !==null) {
				var ret={};
				for (var key in value) {
					if (key.charCodeAt(0) !==charCodeUnderscore) {
						var propValue=value[key];
						if (typeof (propValue)==="object" &&
							propValue !==null &&
							Array.isArray(propValue["items"])) {
							propValue=propValue["items"];
						}
						ret[key]=this.purifyJson(propValue);
					}
				}
				return ret;
			}
			else {
				return value;
			}
		};
		return RetrieveResultImpl;
	}());
	OfficeExtension.RetrieveResultImpl=RetrieveResultImpl;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var Constants=(function () {
		function Constants() {
		}
		Constants.flags="flags";
		Constants.getItemAt="GetItemAt";
		Constants.id="Id";
		Constants.idLowerCase="id";
		Constants.idPrivate="_Id";
		Constants.index="_Index";
		Constants.items="_Items";
		Constants.iterativeExecutor="IterativeExecutor";
		Constants.localDocument="http://document.localhost/";
		Constants.localDocumentApiPrefix="http://document.localhost/_api/";
		Constants.processQuery="ProcessQuery";
		Constants.referenceId="_ReferenceId";
		Constants.isTracked="_IsTracked";
		Constants.sourceLibHeader="SdkVersion";
		Constants.embeddingPageOrigin="EmbeddingPageOrigin";
		Constants.embeddingPageSessionInfo="EmbeddingPageSessionInfo";
		Constants.eventMessageCategory=65536;
		Constants.eventWorkbookId="Workbook";
		Constants.eventSourceRemote="Remote";
		Constants.itemsLowerCase="items";
		Constants.proxy="$proxy";
		Constants.scalarPropertyNames="_scalarPropertyNames";
		Constants.navigationPropertyNames="_navigationPropertyNames";
		Constants.className="_className";
		Constants.isCollection="_isCollection";
		Constants.scalarPropertyUpdateable="_scalarPropertyUpdateable";
		Constants.collectionPropertyPath="_collectionPropertyPath";
		return Constants;
	}());
	OfficeExtension.Constants=Constants;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var versionToken=1;
	var internalConfiguration={
		invokeRequestModifier: function (request) {
			request.DdaMethod.Version=versionToken;
			return request;
		},
		invokeResponseModifier: function (args) {
			versionToken=args.Version;
			if (args.Error) {
				args.error={};
				args.error.Code=args.Error;
			}
			return args;
		}
	};
	var EmbeddedApiStatus;
	(function (EmbeddedApiStatus) {
		EmbeddedApiStatus[EmbeddedApiStatus["Success"]=0]="Success";
		EmbeddedApiStatus[EmbeddedApiStatus["Timeout"]=1]="Timeout";
		EmbeddedApiStatus[EmbeddedApiStatus["InternalError"]=5001]="InternalError";
	})(EmbeddedApiStatus || (EmbeddedApiStatus={}));
	var CommunicationConstants;
	(function (CommunicationConstants) {
		CommunicationConstants.SendingId="sId";
		CommunicationConstants.RespondingId="rId";
		CommunicationConstants.CommandKey="command";
		CommunicationConstants.SessionInfoKey="sessionInfo";
		CommunicationConstants.ParamsKey="params";
		CommunicationConstants.ApiReadyCommand="apiready";
		CommunicationConstants.ExecuteMethodCommand="executeMethod";
		CommunicationConstants.GetAppContextCommand="getAppContext";
		CommunicationConstants.RegisterEventCommand="registerEvent";
		CommunicationConstants.UnregisterEventCommand="unregisterEvent";
		CommunicationConstants.FireEventCommand="fireEvent";
	})(CommunicationConstants || (CommunicationConstants={}));
	var EmbeddedSession=(function (_super) {
		__extends(EmbeddedSession, _super);
		function EmbeddedSession(url, options) {
			_super.call(this);
			this.m_chosenWindow=null;
			this.m_chosenOrigin=null;
			this.m_enabled=true;
			this.m_onMessageHandler=this._onMessage.bind(this);
			this.m_callbackList={};
			this.m_id=0;
			this.m_timeoutId=-1;
			this.m_appContext=null;
			this.m_url=url;
			this.m_options=options;
			if (!this.m_options) {
				this.m_options={ sessionKey: Math.random().toString() };
			}
			if (!this.m_options.sessionKey) {
				this.m_options.sessionKey=Math.random().toString();
			}
			if (!this.m_options.container) {
				this.m_options.container=document.body;
			}
			if (!this.m_options.timeoutInMilliseconds) {
				this.m_options.timeoutInMilliseconds=60000;
			}
			if (!this.m_options.height) {
				this.m_options.height="400px";
			}
			if (!this.m_options.width) {
				this.m_options.width="100%";
			}
		}
		EmbeddedSession.prototype._getIFrameSrc=function () {
			var origin=window.location.protocol+"//"+window.location.host;
			var toAppend=OfficeExtension.Constants.embeddingPageOrigin+"="+encodeURIComponent(origin)+"&"+OfficeExtension.Constants.embeddingPageSessionInfo+"="+encodeURIComponent(this.m_options.sessionKey);
			var useHash=false;
			if (this.m_url.toLowerCase().indexOf("/_layouts/preauth.aspx") > 0 ||
				this.m_url.toLowerCase().indexOf("/_layouts/15/preauth.aspx") > 0) {
				useHash=true;
			}
			var a=document.createElement("a");
			a.href=this.m_url;
			if (useHash) {
				if (a.hash.length===0 || a.hash==="#") {
					a.hash="#"+toAppend;
				}
				else {
					a.hash=a.hash+"&"+toAppend;
				}
			}
			else {
				if (a.search.length===0 || a.search==="?") {
					a.search="?"+toAppend;
				}
				else {
					a.search=a.search+"&"+toAppend;
				}
			}
			var iframeSrc=a.href;
			return iframeSrc;
		};
		EmbeddedSession.prototype.init=function () {
			var _this=this;
			window.addEventListener("message", this.m_onMessageHandler);
			var iframeSrc=this._getIFrameSrc();
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				var iframeElement=document.createElement("iframe");
				if (_this.m_options.id) {
					iframeElement.id=_this.m_options.id;
				}
				iframeElement.style.height=_this.m_options.height;
				iframeElement.style.width=_this.m_options.width;
				iframeElement.src=iframeSrc;
				_this.m_options.container.appendChild(iframeElement);
				_this.m_timeoutId=setTimeout(function () {
					_this.close();
					var err=OfficeExtension.Utility.createRuntimeError(OfficeExtension.ErrorCodes.timeout, OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.timeout), "EmbeddedSession.init");
					reject(err);
				}, _this.m_options.timeoutInMilliseconds);
				_this.m_promiseResolver=resolve;
			});
		};
		EmbeddedSession.prototype._invoke=function (method, callback, params) {
			if (!this.m_enabled) {
				callback(EmbeddedApiStatus.InternalError, null);
				return;
			}
			if (internalConfiguration.invokeRequestModifier) {
				params=internalConfiguration.invokeRequestModifier(params);
			}
			this._sendMessageWithCallback(this.m_id++, method, params, function (args) {
				if (internalConfiguration.invokeResponseModifier) {
					args=internalConfiguration.invokeResponseModifier(args);
				}
				var errorCode=args["Error"];
				delete args["Error"];
				callback(errorCode || EmbeddedApiStatus.Success, args);
			});
		};
		EmbeddedSession.prototype.close=function () {
			window.removeEventListener("message", this.m_onMessageHandler);
			window.clearTimeout(this.m_timeoutId);
			this.m_enabled=false;
		};
		Object.defineProperty(EmbeddedSession.prototype, "eventRegistration", {
			get: function () {
				if (!this.m_sessionEventManager) {
					this.m_sessionEventManager=new OfficeExtension.EventRegistration(this._registerEventImpl.bind(this), this._unregisterEventImpl.bind(this));
				}
				return this.m_sessionEventManager;
			},
			enumerable: true,
			configurable: true
		});
		EmbeddedSession.prototype._createRequestExecutorOrNull=function () {
			return new EmbeddedRequestExecutor(this);
		};
		EmbeddedSession.prototype._resolveRequestUrlAndHeaderInfo=function () {
			return OfficeExtension.Utility._createPromiseFromResult(null);
		};
		EmbeddedSession.prototype._registerEventImpl=function (eventId, targetId) {
			var _this=this;
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				_this._sendMessageWithCallback(_this.m_id++, CommunicationConstants.RegisterEventCommand, { EventId: eventId, TargetId: targetId }, function () {
					resolve(null);
				});
			});
		};
		EmbeddedSession.prototype._unregisterEventImpl=function (eventId, targetId) {
			var _this=this;
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				_this._sendMessageWithCallback(_this.m_id++, CommunicationConstants.UnregisterEventCommand, { EventId: eventId, TargetId: targetId }, function () {
					resolve();
				});
			});
		};
		EmbeddedSession.prototype._onMessage=function (event) {
			var _this=this;
			if (!this.m_enabled) {
				return;
			}
			if (this.m_chosenWindow
				&& (this.m_chosenWindow !==event.source || this.m_chosenOrigin !==event.origin)) {
				return;
			}
			var eventData=event.data;
			if (eventData && eventData[CommunicationConstants.CommandKey]===CommunicationConstants.ApiReadyCommand) {
				if (!this.m_chosenWindow
					&& this._isValidDescendant(event.source)
					&& eventData[CommunicationConstants.SessionInfoKey]===this.m_options.sessionKey) {
					this.m_chosenWindow=event.source;
					this.m_chosenOrigin=event.origin;
					this._sendMessageWithCallback(this.m_id++, CommunicationConstants.GetAppContextCommand, null, function (appContext) {
						_this._setupContext(appContext);
						window.clearTimeout(_this.m_timeoutId);
						_this.m_promiseResolver();
					});
				}
				return;
			}
			if (eventData && eventData[CommunicationConstants.CommandKey]===CommunicationConstants.FireEventCommand) {
				var msg=eventData[CommunicationConstants.ParamsKey];
				var eventId=msg["EventId"];
				var targetId=msg["TargetId"];
				var data=msg["Data"];
				if (this.m_sessionEventManager) {
					var handlers=this.m_sessionEventManager.getHandlers(eventId, targetId);
					for (var i=0; i < handlers.length; i++) {
						handlers[i](data);
					}
				}
				return;
			}
			if (eventData && eventData.hasOwnProperty(CommunicationConstants.RespondingId)) {
				var rId=eventData[CommunicationConstants.RespondingId];
				var callback=this.m_callbackList[rId];
				if (typeof callback==="function") {
					callback(eventData[CommunicationConstants.ParamsKey]);
				}
				delete this.m_callbackList[rId];
			}
		};
		EmbeddedSession.prototype._sendMessageWithCallback=function (id, command, data, callback) {
			this.m_callbackList[id]=callback;
			var message={};
			message[CommunicationConstants.SendingId]=id;
			message[CommunicationConstants.CommandKey]=command;
			message[CommunicationConstants.ParamsKey]=data;
			this.m_chosenWindow.postMessage(JSON.stringify(message), this.m_chosenOrigin);
		};
		EmbeddedSession.prototype._isValidDescendant=function (wnd) {
			var container=this.m_options.container || document.body;
			function doesFrameWindow(containerWindow) {
				if (containerWindow===wnd) {
					return true;
				}
				for (var i=0, len=containerWindow.frames.length; i < len; i++) {
					if (doesFrameWindow(containerWindow.frames[i])) {
						return true;
					}
				}
				return false;
			}
			var iframes=container.getElementsByTagName("iframe");
			for (var i=0, len=iframes.length; i < len; i++) {
				if (doesFrameWindow(iframes[i].contentWindow)) {
					return true;
				}
			}
			return false;
		};
		EmbeddedSession.prototype._setupContext=function (appContext) {
			if (!(this.m_appContext=appContext)) {
				return;
			}
		};
		return EmbeddedSession;
	}(OfficeExtension.SessionBase));
	OfficeExtension.EmbeddedSession=EmbeddedSession;
	var EmbeddedRequestExecutor=(function () {
		function EmbeddedRequestExecutor(session) {
			this.m_session=session;
		}
		EmbeddedRequestExecutor.prototype.executeAsync=function (customData, requestFlags, requestMessage) {
			var _this=this;
			var messageSafearray=OfficeExtension.RichApiMessageUtility.buildMessageArrayForIRequestExecutor(customData, requestFlags, requestMessage, EmbeddedRequestExecutor.SourceLibHeaderValue);
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				_this.m_session._invoke(CommunicationConstants.ExecuteMethodCommand, function (status, result) {
					OfficeExtension.Utility.log("Response:");
					OfficeExtension.Utility.log(JSON.stringify(result));
					var response;
					if (status==EmbeddedApiStatus.Success) {
						response=OfficeExtension.RichApiMessageUtility.buildResponseOnSuccess(OfficeExtension.RichApiMessageUtility.getResponseBodyFromSafeArray(result.Data), OfficeExtension.RichApiMessageUtility.getResponseHeadersFromSafeArray(result.Data));
					}
					else {
						response=OfficeExtension.RichApiMessageUtility.buildResponseOnError(result.error.Code, result.error.Message);
					}
					resolve(response);
				}, EmbeddedRequestExecutor._transformMessageArrayIntoParams(messageSafearray));
			});
		};
		EmbeddedRequestExecutor._transformMessageArrayIntoParams=function (msgArray) {
			return {
				ArrayData: msgArray,
				DdaMethod: {
					DispatchId: EmbeddedRequestExecutor.DispidExecuteRichApiRequestMethod
				}
			};
		};
		EmbeddedRequestExecutor.DispidExecuteRichApiRequestMethod=93;
		EmbeddedRequestExecutor.SourceLibHeaderValue="Embedded";
		return EmbeddedRequestExecutor;
	}());
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var _Internal;
	(function (_Internal) {
		var RuntimeError=(function (_super) {
			__extends(RuntimeError, _super);
			function RuntimeError(error) {
				_super.call(this, (typeof error==="string") ? error : error.message);
				this.name="OfficeExtension.Error";
				if (typeof error==="string") {
					this.message=error;
				}
				else {
					this.code=error.code;
					this.message=error.message;
					this.traceMessages=error.traceMessages || [];
					this.innerError=error.innerError || null;
					this.debugInfo=this._createDebugInfo(error.debugInfo || {});
				}
			}
			RuntimeError.prototype.toString=function () {
				return this.code+': '+this.message;
			};
			RuntimeError.prototype._createDebugInfo=function (partialDebugInfo) {
				var debugInfo={
					code: this.code,
					message: this.message,
					toString: function () {
						return JSON.stringify(this);
					}
				};
				for (var key in partialDebugInfo) {
					debugInfo[key]=partialDebugInfo[key];
				}
				if (this.innerError) {
					if (this.innerError instanceof OfficeExtension._Internal.RuntimeError) {
						debugInfo.innerError=this.innerError.debugInfo;
					}
					else {
						debugInfo.innerError=this.innerError;
					}
				}
				return debugInfo;
			};
			RuntimeError._createInvalidArgError=function (error) {
				return new _Internal.RuntimeError({
					code: OfficeExtension.ErrorCodes.invalidArgument,
					message: (OfficeExtension.Utility.isNullOrEmptyString(error.argumentName) ?
						OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.invalidArgumentGeneric) :
						OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.invalidArgument, error.argumentName)),
					debugInfo: error.errorLocation ? { errorLocation: error.errorLocation } : {},
					innerError: error.innerError
				});
			};
			return RuntimeError;
		}(Error));
		_Internal.RuntimeError=RuntimeError;
	})(_Internal=OfficeExtension._Internal || (OfficeExtension._Internal={}));
	OfficeExtension.Error=_Internal.RuntimeError;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ErrorCodes=(function () {
		function ErrorCodes() {
		}
		ErrorCodes.accessDenied="AccessDenied";
		ErrorCodes.generalException="GeneralException";
		ErrorCodes.activityLimitReached="ActivityLimitReached";
		ErrorCodes.invalidObjectPath="InvalidObjectPath";
		ErrorCodes.propertyNotLoaded="PropertyNotLoaded";
		ErrorCodes.valueNotLoaded="ValueNotLoaded";
		ErrorCodes.invalidRequestContext="InvalidRequestContext";
		ErrorCodes.invalidArgument="InvalidArgument";
		ErrorCodes.runMustReturnPromise="RunMustReturnPromise";
		ErrorCodes.cannotRegisterEvent="CannotRegisterEvent";
		ErrorCodes.apiNotFound="ApiNotFound";
		ErrorCodes.connectionFailure="ConnectionFailure";
		ErrorCodes.timeout="Timeout";
		ErrorCodes.invalidOrTimedOutSession="InvalidOrTimedOutSession";
		ErrorCodes.cannotUpdateReadOnlyProperty="CannotUpdateReadOnlyProperty";
		return ErrorCodes;
	}());
	OfficeExtension.ErrorCodes=ErrorCodes;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var EventHandlers=(function () {
		function EventHandlers(context, parentObject, name, eventInfo) {
			var _this=this;
			this.m_id=context._nextId();
			this.m_context=context;
			this.m_name=name;
			this.m_handlers=[];
			this.m_registered=false;
			this.m_eventInfo=eventInfo;
			this.m_callback=function (args) {
				_this.m_eventInfo.eventArgsTransformFunc(args)
					.then(function (newArgs) { return _this.fireEvent(newArgs); });
			};
		}
		Object.defineProperty(EventHandlers.prototype, "_registered", {
			get: function () {
				return this.m_registered;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(EventHandlers.prototype, "_id", {
			get: function () {
				return this.m_id;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(EventHandlers.prototype, "_handlers", {
			get: function () {
				return this.m_handlers;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(EventHandlers.prototype, "_context", {
			get: function () {
				return this.m_context;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(EventHandlers.prototype, "_callback", {
			get: function () {
				return this.m_callback;
			},
			enumerable: true,
			configurable: true
		});
		EventHandlers.prototype.add=function (handler) {
			var action=OfficeExtension.ActionFactory.createTraceAction(this.m_context, null, false);
			this.m_context._pendingRequest._addPendingEventHandlerAction(this, { id: action.actionInfo.Id, handler: handler, operation: 0 });
			return new OfficeExtension.EventHandlerResult(this.m_context, this, handler);
		};
		EventHandlers.prototype.remove=function (handler) {
			var action=OfficeExtension.ActionFactory.createTraceAction(this.m_context, null, false);
			this.m_context._pendingRequest._addPendingEventHandlerAction(this, { id: action.actionInfo.Id, handler: handler, operation: 1 });
		};
		EventHandlers.prototype.removeAll=function () {
			var action=OfficeExtension.ActionFactory.createTraceAction(this.m_context, null, false);
			this.m_context._pendingRequest._addPendingEventHandlerAction(this, { id: action.actionInfo.Id, handler: null, operation: 2 });
		};
		EventHandlers.prototype._processRegistration=function (req) {
			var _this=this;
			var ret=OfficeExtension.Utility._createPromiseFromResult(null);
			var actions=req._getPendingEventHandlerActions(this);
			if (!actions) {
				return ret;
			}
			var handlersResult=[];
			for (var i=0; i < this.m_handlers.length; i++) {
				handlersResult.push(this.m_handlers[i]);
			}
			var hasChange=false;
			for (var i=0; i < actions.length; i++) {
				if (req._responseTraceIds[actions[i].id]) {
					hasChange=true;
					switch (actions[i].operation) {
						case 0:
							handlersResult.push(actions[i].handler);
							break;
						case 1:
							for (var index=handlersResult.length - 1; index >=0; index--) {
								if (handlersResult[index]===actions[i].handler) {
									handlersResult.splice(index, 1);
									break;
								}
							}
							break;
						case 2:
							handlersResult=[];
							break;
					}
				}
			}
			if (hasChange) {
				if (!this.m_registered && handlersResult.length > 0) {
					ret=ret
						.then(function () { return _this.m_eventInfo.registerFunc(_this.m_callback); })
						.then(function () { return (_this.m_registered=true); });
				}
				else if (this.m_registered && handlersResult.length==0) {
					ret=ret
						.then(function () { return _this.m_eventInfo.unregisterFunc(_this.m_callback); })
						.catch(function (ex) {
						OfficeExtension.Utility.log("Error when unregister event: "+JSON.stringify(ex));
					})
						.then(function () { return (_this.m_registered=false); });
				}
				ret=ret
					.then(function () { return (_this.m_handlers=handlersResult); });
			}
			return ret;
		};
		EventHandlers.prototype.fireEvent=function (args) {
			var promises=[];
			for (var i=0; i < this.m_handlers.length; i++) {
				var handler=this.m_handlers[i];
				var p=OfficeExtension.Utility._createPromiseFromResult(null)
					.then(this.createFireOneEventHandlerFunc(handler, args))
					.catch(function (ex) {
					OfficeExtension.Utility.log("Error when invoke handler: "+JSON.stringify(ex));
				});
				promises.push(p);
			}
			OfficeExtension._Internal.OfficePromise.all(promises);
		};
		EventHandlers.prototype.createFireOneEventHandlerFunc=function (handler, args) {
			return function () { return handler(args); };
		};
		return EventHandlers;
	}());
	OfficeExtension.EventHandlers=EventHandlers;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var EventHandlerResult=(function () {
		function EventHandlerResult(context, handlers, handler) {
			this.m_context=context;
			this.m_allHandlers=handlers;
			this.m_handler=handler;
		}
		Object.defineProperty(EventHandlerResult.prototype, "context", {
			get: function () {
				return this.m_context;
			},
			enumerable: true,
			configurable: true
		});
		EventHandlerResult.prototype.remove=function () {
			if (this.m_allHandlers && this.m_handler) {
				this.m_allHandlers.remove(this.m_handler);
				this.m_allHandlers=null;
				this.m_handler=null;
			}
		};
		return EventHandlerResult;
	}());
	OfficeExtension.EventHandlerResult=EventHandlerResult;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var _Internal;
	(function (_Internal) {
		var OfficeJsEventRegistration=(function () {
			function OfficeJsEventRegistration() {
			}
			OfficeJsEventRegistration.prototype.register=function (eventId, targetId, handler) {
				switch (eventId) {
					case 4:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(targetId, callback); })
							.then(function (officeBinding) {
							return OfficeExtension.Utility.promisify(function (callback) { return officeBinding.addHandlerAsync(Office.EventType.BindingDataChanged, handler, callback); });
						});
					case 3:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(targetId, callback); })
							.then(function (officeBinding) {
							return OfficeExtension.Utility.promisify(function (callback) { return officeBinding.addHandlerAsync(Office.EventType.BindingSelectionChanged, handler, callback); });
						});
					case 2:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handler, callback); });
					case 1:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.settings.addHandlerAsync(Office.EventType.SettingsChanged, handler, callback); });
					case 5:
						return OfficeExtension.Utility.promisify(function (callback) { return OSF.DDA.RichApi.richApiMessageManager.addHandlerAsync("richApiMessage", handler, callback); });
					case 13:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.addHandlerAsync(Office.EventType.ObjectDeleted, handler, { id: targetId }, callback); });
					case 14:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.addHandlerAsync(Office.EventType.ObjectSelectionChanged, handler, { id: targetId }, callback); });
					case 15:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.addHandlerAsync(Office.EventType.ObjectDataChanged, handler, { id: targetId }, callback); });
					case 16:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.addHandlerAsync(Office.EventType.ContentControlAdded, handler, { id: targetId }, callback); });
					default:
						throw _Internal.RuntimeError._createInvalidArgError("eventId");
				}
			};
			OfficeJsEventRegistration.prototype.unregister=function (eventId, targetId, handler) {
				switch (eventId) {
					case 4:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(targetId, callback); })
							.then(function (officeBinding) {
							return OfficeExtension.Utility.promisify(function (callback) { return officeBinding.removeHandlerAsync(Office.EventType.BindingDataChanged, { handler: handler }, callback); });
						});
					case 3:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(targetId, callback); })
							.then(function (officeBinding) {
							return OfficeExtension.Utility.promisify(function (callback) { return officeBinding.removeHandlerAsync(Office.EventType.BindingSelectionChanged, { handler: handler }, callback); });
						});
					case 2:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, { handler: handler }, callback); });
					case 1:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.settings.removeHandlerAsync(Office.EventType.SettingsChanged, { handler: handler }, callback); });
					case 5:
						return OfficeExtension.Utility.promisify(function (callback) { return OSF.DDA.RichApi.richApiMessageManager.removeHandlerAsync("richApiMessage", { handler: handler }, callback); });
					case 13:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.removeHandlerAsync(Office.EventType.ObjectDeleted, { id: targetId, handler: handler }, callback); });
					case 14:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.removeHandlerAsync(Office.EventType.ObjectSelectionChanged, { id: targetId, handler: handler }, callback); });
					case 15:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.removeHandlerAsync(Office.EventType.ObjectDataChanged, { id: targetId, handler: handler }, callback); });
					case 16:
						return OfficeExtension.Utility.promisify(function (callback) { return Office.context.document.removeHandlerAsync(Office.EventType.ContentControlAdded, { id: targetId, handler: handler }, callback); });
					default:
						throw _Internal.RuntimeError._createInvalidArgError("eventId");
				}
			};
			return OfficeJsEventRegistration;
		}());
		_Internal.officeJsEventRegistration=new OfficeJsEventRegistration();
	})(_Internal=OfficeExtension._Internal || (OfficeExtension._Internal={}));
	var EventRegistration=(function () {
		function EventRegistration(registerEventImpl, unregisterEventImpl) {
			this.m_handlersByEventByTarget={};
			this.m_registerEventImpl=registerEventImpl;
			this.m_unregisterEventImpl=unregisterEventImpl;
		}
		EventRegistration.prototype.getHandlers=function (eventId, targetId) {
			if (OfficeExtension.Utility.isNullOrUndefined(targetId)) {
				targetId="";
			}
			var handlersById=this.m_handlersByEventByTarget[eventId];
			if (!handlersById) {
				handlersById={};
				this.m_handlersByEventByTarget[eventId]=handlersById;
			}
			var handlers=handlersById[targetId];
			if (!handlers) {
				handlers=[];
				handlersById[targetId]=handlers;
			}
			return handlers;
		};
		EventRegistration.prototype.register=function (eventId, targetId, handler) {
			if (!handler) {
				throw _Internal.RuntimeError._createInvalidArgError("handler");
			}
			var handlers=this.getHandlers(eventId, targetId);
			handlers.push(handler);
			if (handlers.length===1) {
				return this.m_registerEventImpl(eventId, targetId);
			}
			return OfficeExtension.Utility._createPromiseFromResult(null);
		};
		EventRegistration.prototype.unregister=function (eventId, targetId, handler) {
			if (!handler) {
				throw _Internal.RuntimeError._createInvalidArgError("handler");
			}
			var handlers=this.getHandlers(eventId, targetId);
			for (var index=handlers.length - 1; index >=0; index--) {
				if (handlers[index]===handler) {
					handlers.splice(index, 1);
					break;
				}
			}
			if (handlers.length===0) {
				return this.m_unregisterEventImpl(eventId, targetId);
			}
			return OfficeExtension.Utility._createPromiseFromResult(null);
		};
		return EventRegistration;
	}());
	OfficeExtension.EventRegistration=EventRegistration;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var GenericEventRegistration=(function () {
		function GenericEventRegistration() {
			this.m_eventRegistration=new OfficeExtension.EventRegistration(this._registerEventImpl.bind(this), this._unregisterEventImpl.bind(this));
			this.m_richApiMessageHandler=this._handleRichApiMessage.bind(this);
		}
		GenericEventRegistration.prototype.ready=function () {
			var _this=this;
			if (!this.m_ready) {
				if (GenericEventRegistration._testReadyImpl) {
					this.m_ready=GenericEventRegistration._testReadyImpl()
						.then(function () {
						_this.m_isReady=true;
					});
				}
				else {
					this.m_ready=OfficeExtension._Internal.officeJsEventRegistration.register(5, "", this.m_richApiMessageHandler)
						.then(function () {
						_this.m_isReady=true;
					});
				}
			}
			return this.m_ready;
		};
		Object.defineProperty(GenericEventRegistration.prototype, "isReady", {
			get: function () {
				return this.m_isReady;
			},
			enumerable: true,
			configurable: true
		});
		GenericEventRegistration.prototype.register=function (eventId, targetId, handler) {
			var _this=this;
			return this.ready()
				.then(function () { return _this.m_eventRegistration.register(eventId, targetId, handler); });
		};
		GenericEventRegistration.prototype.unregister=function (eventId, targetId, handler) {
			var _this=this;
			return this.ready()
				.then(function () { return _this.m_eventRegistration.unregister(eventId, targetId, handler); });
		};
		GenericEventRegistration.prototype._registerEventImpl=function (eventId, targetId) {
			return OfficeExtension.Utility._createPromiseFromResult(null);
		};
		GenericEventRegistration.prototype._unregisterEventImpl=function (eventId, targetId) {
			return OfficeExtension.Utility._createPromiseFromResult(null);
		};
		GenericEventRegistration.prototype._handleRichApiMessage=function (msg) {
			if (msg && msg.entries) {
				for (var entryIndex=0; entryIndex < msg.entries.length; entryIndex++) {
					var entry=msg.entries[entryIndex];
					if (entry.messageCategory==OfficeExtension.Constants.eventMessageCategory) {
						if (OfficeExtension.Utility._logEnabled) {
							OfficeExtension.Utility.log(JSON.stringify(entry));
						}
						var funcs=this.m_eventRegistration.getHandlers(entry.messageType, entry.targetId);
						if (funcs.length > 0) {
							var arg=JSON.parse(entry.message);
							if (entry.isRemoteOverride) {
								arg.source=OfficeExtension.Constants.eventSourceRemote;
							}
							for (var i=0; i < funcs.length; i++) {
								funcs[i](arg);
							}
						}
					}
				}
			}
		};
		GenericEventRegistration.getGenericEventRegistration=function () {
			if (!GenericEventRegistration.s_genericEventRegistration) {
				GenericEventRegistration.s_genericEventRegistration=new GenericEventRegistration();
			}
			return GenericEventRegistration.s_genericEventRegistration;
		};
		GenericEventRegistration.richApiMessageEventCategory=65536;
		return GenericEventRegistration;
	}());
	function _testSetRichApiMessageReadyImpl(impl) {
		GenericEventRegistration._testReadyImpl=impl;
	}
	OfficeExtension._testSetRichApiMessageReadyImpl=_testSetRichApiMessageReadyImpl;
	function _testTriggerRichApiMessageEvent(msg) {
		GenericEventRegistration.getGenericEventRegistration()._handleRichApiMessage(msg);
	}
	OfficeExtension._testTriggerRichApiMessageEvent=_testTriggerRichApiMessageEvent;
	var GenericEventHandlers=(function (_super) {
		__extends(GenericEventHandlers, _super);
		function GenericEventHandlers(context, parentObject, name, eventInfo) {
			_super.call(this, context, parentObject, name, eventInfo);
			this.m_genericEventInfo=eventInfo;
		}
		GenericEventHandlers.prototype.add=function (handler) {
			var _this=this;
			if ((this._handlers.length==0) && this.m_genericEventInfo.registerFunc) {
				this.m_genericEventInfo.registerFunc();
			}
			if (!GenericEventRegistration.getGenericEventRegistration().isReady) {
				this._context._pendingRequest._addPreSyncPromise(GenericEventRegistration.getGenericEventRegistration().ready());
			}
			OfficeExtension.ActionFactory.createTraceMarkerForCallback(this._context, function () {
				_this._handlers.push(handler);
				if (_this._handlers.length==1) {
					GenericEventRegistration.getGenericEventRegistration().register(_this.m_genericEventInfo.eventType, _this.m_genericEventInfo.getTargetIdFunc(), _this._callback);
				}
			});
			return new OfficeExtension.EventHandlerResult(this._context, this, handler);
		};
		GenericEventHandlers.prototype.remove=function (handler) {
			var _this=this;
			if ((this._handlers.length==1) && this.m_genericEventInfo.unregisterFunc) {
				this.m_genericEventInfo.unregisterFunc();
			}
			OfficeExtension.ActionFactory.createTraceMarkerForCallback(this._context, function () {
				var handlers=_this._handlers;
				for (var index=handlers.length - 1; index >=0; index--) {
					if (handlers[index]===handler) {
						handlers.splice(index, 1);
						break;
					}
				}
				if (handlers.length==0) {
					GenericEventRegistration.getGenericEventRegistration().unregister(_this.m_genericEventInfo.eventType, _this.m_genericEventInfo.getTargetIdFunc(), _this._callback);
				}
			});
		};
		GenericEventHandlers.prototype.removeAll=function () {
		};
		return GenericEventHandlers;
	}(OfficeExtension.EventHandlers));
	OfficeExtension.GenericEventHandlers=GenericEventHandlers;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var HttpRequestExecutor=(function () {
		function HttpRequestExecutor() {
		}
		HttpRequestExecutor.prototype.executeAsync=function (customData, requestFlags, requestMessage) {
			var requestMessageText=JSON.stringify(requestMessage.Body);
			var url=requestMessage.Url;
			if (url.charAt(url.length - 1) !="/") {
				url=url+"/";
			}
			url=url+OfficeExtension.Constants.processQuery;
			url=url+"?"+OfficeExtension.Constants.flags+"="+requestFlags.toString();
			var requestInfo={
				method: "POST",
				url: url,
				headers: {},
				body: requestMessageText
			};
			requestInfo.headers[OfficeExtension.Constants.sourceLibHeader]=HttpRequestExecutor.SourceLibHeaderValue;
			requestInfo.headers["CONTENT-TYPE"]="application/json";
			if (requestMessage.Headers) {
				for (var key in requestMessage.Headers) {
					requestInfo.headers[key]=requestMessage.Headers[key];
				}
			}
			return OfficeExtension.HttpUtility.sendRequest(requestInfo)
				.then(function (responseInfo) {
				var response;
				if (responseInfo.statusCode===200) {
					response={ ErrorCode: null, ErrorMessage: null, Headers: responseInfo.headers, Body: JSON.parse(responseInfo.body) };
				}
				else {
					OfficeExtension.Utility.log("Error Response:"+responseInfo.body);
					var error=OfficeExtension.Utility._parseErrorResponse(responseInfo);
					response={
						ErrorCode: error.errorCode,
						ErrorMessage: error.errorMessage,
						Headers: responseInfo.headers,
						Body: null
					};
				}
				return response;
			});
		};
		HttpRequestExecutor.SourceLibHeaderValue="officejs-rest";
		return HttpRequestExecutor;
	}());
	OfficeExtension.HttpRequestExecutor=HttpRequestExecutor;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var HttpUtility=(function () {
		function HttpUtility() {
		}
		HttpUtility.setCustomSendRequestFunc=function (func) {
			HttpUtility.s_customSendRequestFunc=func;
		};
		HttpUtility.xhrSendRequestFunc=function (request) {
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				var xhr=new XMLHttpRequest();
				xhr.open(request.method, request.url);
				xhr.onload=function () {
					var resp={
						statusCode: xhr.status,
						headers: OfficeExtension.Utility._parseHttpResponseHeaders(xhr.getAllResponseHeaders()),
						body: xhr.responseText
					};
					resolve(resp);
				};
				xhr.onerror=function () {
					reject(new OfficeExtension._Internal.RuntimeError({
						code: OfficeExtension.ErrorCodes.connectionFailure,
						message: OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.connectionFailureWithStatus, xhr.statusText)
					}));
				};
				if (request.headers) {
					for (var key in request.headers) {
						xhr.setRequestHeader(key, request.headers[key]);
					}
				}
				xhr.send(request.body);
			});
		};
		HttpUtility.sendRequest=function (request) {
			HttpUtility.validateAndNormalizeRequest(request);
			var func=HttpUtility.s_customSendRequestFunc;
			if (!func) {
				func=HttpUtility.xhrSendRequestFunc;
			}
			return func(request);
		};
		HttpUtility.setCustomSendLocalDocumentRequestFunc=function (func) {
			HttpUtility.s_customSendLocalDocumentRequestFunc=func;
		};
		HttpUtility.sendLocalDocumentRequest=function (request) {
			HttpUtility.validateAndNormalizeRequest(request);
			var func;
			func=HttpUtility.s_customSendLocalDocumentRequestFunc || HttpUtility.officeJsSendLocalDocumentRequestFunc;
			return func(request);
		};
		HttpUtility.officeJsSendLocalDocumentRequestFunc=function (request) {
			request=OfficeExtension.Utility._validateLocalDocumentRequest(request);
			var requestSafeArray=OfficeExtension.Utility._buildRequestMessageSafeArray(request);
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				OSF.DDA.RichApi.executeRichApiRequestAsync(requestSafeArray, function (asyncResult) {
					var response;
					if (asyncResult.status=="succeeded") {
						response=							{
								statusCode: OfficeExtension.RichApiMessageUtility.getResponseStatusCode(asyncResult),
								headers: OfficeExtension.RichApiMessageUtility.getResponseHeaders(asyncResult),
								body: OfficeExtension.RichApiMessageUtility.getResponseBody(asyncResult)
							};
					}
					else {
						response=OfficeExtension.RichApiMessageUtility.buildHttpResponseFromOfficeJsError(asyncResult.error.code, asyncResult.error.message);
					}
					OfficeExtension.Utility.log(JSON.stringify(response));
					resolve(response);
				});
			});
		};
		HttpUtility.validateAndNormalizeRequest=function (request) {
			if (OfficeExtension.Utility.isNullOrUndefined(request)) {
				throw OfficeExtension._Internal.RuntimeError._createInvalidArgError({
					argumentName: "request"
				});
			}
			if (OfficeExtension.Utility.isNullOrEmptyString(request.method)) {
				request.method="GET";
			}
			request.method=request.method.toUpperCase();
		};
		HttpUtility.logRequest=function (request) {
			if (OfficeExtension.Utility._logEnabled) {
				OfficeExtension.Utility.log("---HTTP Request---");
				OfficeExtension.Utility.log(request.method+" "+request.url);
				if (request.headers) {
					for (var key in request.headers) {
						OfficeExtension.Utility.log(key+": "+request.headers[key]);
					}
				}
				if (HttpUtility._logBody) {
					OfficeExtension.Utility.log(request.body);
				}
			}
		};
		HttpUtility.logResponse=function (response) {
			if (OfficeExtension.Utility._logEnabled) {
				OfficeExtension.Utility.log("---HTTP Response---");
				OfficeExtension.Utility.log(""+response.statusCode);
				if (response.headers) {
					for (var key in response.headers) {
						OfficeExtension.Utility.log(key+": "+response.headers[key]);
					}
				}
				if (HttpUtility._logBody) {
					OfficeExtension.Utility.log(response.body);
				}
			}
		};
		HttpUtility._logBody=false;
		return HttpUtility;
	}());
	OfficeExtension.HttpUtility=HttpUtility;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var InstantiateActionResultHandler=(function () {
		function InstantiateActionResultHandler(clientObject) {
			this.m_clientObject=clientObject;
		}
		InstantiateActionResultHandler.prototype._handleResult=function (value) {
			this.m_clientObject._handleIdResult(value);
		};
		return InstantiateActionResultHandler;
	}());
	OfficeExtension.InstantiateActionResultHandler=InstantiateActionResultHandler;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var HostBridgeRequestExecutor=(function () {
		function HostBridgeRequestExecutor(session) {
			this.m_session=session;
		}
		HostBridgeRequestExecutor.prototype.executeAsync=function (customData, requestFlags, requestMessage) {
			var httpRequestInfo={
				url: OfficeExtension.Constants.processQuery,
				method: "POST",
				headers: requestMessage.Headers,
				body: requestMessage.Body
			};
			var message={
				id: HostBridgeSession.nextId(),
				type: 1,
				flags: requestFlags,
				message: httpRequestInfo
			};
			OfficeExtension.Utility.log(JSON.stringify(message));
			return this.m_session.sendMessageToHost(message)
				.then(function (nativeBridgeResponse) {
				OfficeExtension.Utility.log("Received response: "+JSON.stringify(nativeBridgeResponse));
				var responseInfo=nativeBridgeResponse.message;
				var response;
				if (responseInfo.statusCode===200) {
					response={ ErrorCode: null, ErrorMessage: null, Headers: responseInfo.headers, Body: responseInfo.body };
				}
				else {
					OfficeExtension.Utility.log("Error Response:"+responseInfo.body);
					var error=OfficeExtension.Utility._parseErrorResponse(responseInfo);
					response={
						ErrorCode: error.errorCode,
						ErrorMessage: error.errorMessage,
						Headers: responseInfo.headers,
						Body: null
					};
				}
				return response;
			});
		};
		return HostBridgeRequestExecutor;
	}());
	var HostBridgeSession=(function (_super) {
		__extends(HostBridgeSession, _super);
		function HostBridgeSession(bridge) {
			var _this=this;
			_super.call(this);
			this.m_promiseResolver={};
			this.m_bridge=bridge;
			this.m_bridge.onMessageFromHost=function (msg) {
				_this.onMessageFromHost(msg);
			};
		}
		HostBridgeSession.prototype._resolveRequestUrlAndHeaderInfo=function () {
			return OfficeExtension.Utility._createPromiseFromResult(null);
		};
		HostBridgeSession.prototype._createRequestExecutorOrNull=function () {
			OfficeExtension.Utility.log("NativeBridgeSession::CreateRequestExecutor");
			return new HostBridgeRequestExecutor(this);
		};
		Object.defineProperty(HostBridgeSession.prototype, "eventRegistration", {
			get: function () {
				return OfficeExtension._Internal.officeJsEventRegistration;
			},
			enumerable: true,
			configurable: true
		});
		HostBridgeSession.init=function (bridge) {
			if (bridge && typeof (bridge)==="object") {
				var session=new HostBridgeSession(bridge);
				OfficeExtension.ClientRequestContext._overrideSession=session;
				OfficeExtension.HttpUtility.setCustomSendLocalDocumentRequestFunc(function (request) {
					var bridgeMessage={
						id: HostBridgeSession.nextId(),
						type: 1,
						flags: 0,
						message: request
					};
					return session.sendMessageToHost(bridgeMessage)
						.then(function (bridgeResponse) {
						var responseInfo=bridgeResponse.message;
						return responseInfo;
					});
				});
			}
		};
		HostBridgeSession.prototype.sendMessageToHost=function (message) {
			var _this=this;
			this.m_bridge.sendMessageToHost(JSON.stringify(message));
			var ret=new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				_this.m_promiseResolver[message.id]=resolve;
			});
			return ret;
		};
		HostBridgeSession.prototype.onMessageFromHost=function (messageText) {
			if (messageText==="test") {
				if (HostBridgeTest._testFunc) {
					HostBridgeTest._testFunc();
				}
			}
			else {
				var message=JSON.parse(messageText);
				if (typeof (message.id)==="number") {
					var resolve=this.m_promiseResolver[message.id];
					if (resolve) {
						resolve(message);
					}
					delete this.m_promiseResolver[message.id];
				}
			}
		};
		HostBridgeSession.nextId=function () {
			return HostBridgeSession.s_nextId++;
		};
		HostBridgeSession.s_nextId=1;
		return HostBridgeSession;
	}(OfficeExtension.SessionBase));
	var HostBridge=(function () {
		function HostBridge() {
		}
		HostBridge.init=function (bridge) {
			HostBridgeSession.init(bridge);
		};
		return HostBridge;
	}());
	OfficeExtension.HostBridge=HostBridge;
	if (typeof (_richApiNativeBridge)==="object" && _richApiNativeBridge) {
		HostBridge.init(_richApiNativeBridge);
	}
	var HostBridgeTest=(function () {
		function HostBridgeTest() {
		}
		HostBridgeTest.setTestFunc=function (func) {
			HostBridgeTest._testFunc=func;
		};
		return HostBridgeTest;
	}());
	OfficeExtension.HostBridgeTest=HostBridgeTest;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ObjectPath=(function () {
		function ObjectPath(objectPathInfo, parentObjectPath, isCollection, isInvalidAfterRequest) {
			this.m_objectPathInfo=objectPathInfo;
			this.m_parentObjectPath=parentObjectPath;
			this.m_isWriteOperation=false;
			this.m_isCollection=isCollection;
			this.m_isInvalidAfterRequest=isInvalidAfterRequest;
			this.m_isValid=true;
		}
		Object.defineProperty(ObjectPath.prototype, "objectPathInfo", {
			get: function () {
				return this.m_objectPathInfo;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "isWriteOperation", {
			get: function () {
				return this.m_isWriteOperation;
			},
			set: function (value) {
				this.m_isWriteOperation=value;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "isRestrictedResourceAccess", {
			get: function () {
				return this.m_isRestrictedResourceAccess;
			},
			set: function (value) {
				this.m_isRestrictedResourceAccess=value;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "isCollection", {
			get: function () {
				return this.m_isCollection;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "isInvalidAfterRequest", {
			get: function () {
				return this.m_isInvalidAfterRequest;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "parentObjectPath", {
			get: function () {
				return this.m_parentObjectPath;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "argumentObjectPaths", {
			get: function () {
				return this.m_argumentObjectPaths;
			},
			set: function (value) {
				this.m_argumentObjectPaths=value;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "isValid", {
			get: function () {
				return this.m_isValid;
			},
			set: function (value) {
				this.m_isValid=value;
				if (!value &&
					this.m_objectPathInfo.ObjectPathType===6 &&
					this.m_savedObjectPathInfo) {
					ObjectPath.copyObjectPathInfo(this.m_savedObjectPathInfo.pathInfo, this.m_objectPathInfo);
					this.m_parentObjectPath=this.m_savedObjectPathInfo.parent;
					this.m_isValid=true;
					this.m_savedObjectPathInfo=null;
				}
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ObjectPath.prototype, "getByIdMethodName", {
			get: function () {
				return this.m_getByIdMethodName;
			},
			set: function (value) {
				this.m_getByIdMethodName=value;
			},
			enumerable: true,
			configurable: true
		});
		ObjectPath.prototype._updateAsNullObject=function () {
			this.resetForUpdateUsingObjectData();
			this.m_objectPathInfo.ObjectPathType=7;
			this.m_objectPathInfo.Name="";
			this.m_parentObjectPath=null;
		};
		ObjectPath.prototype.updateUsingObjectData=function (value, clientObject) {
			var referenceId=value[OfficeExtension.Constants.referenceId];
			if (!OfficeExtension.Utility.isNullOrEmptyString(referenceId)) {
				if (!this.m_savedObjectPathInfo &&
					!this.isInvalidAfterRequest &&
					ObjectPath.isRestorableObjectPath(this.m_objectPathInfo.ObjectPathType)) {
					var pathInfo={};
					ObjectPath.copyObjectPathInfo(this.m_objectPathInfo, pathInfo);
					this.m_savedObjectPathInfo={
						pathInfo: pathInfo,
						parent: this.m_parentObjectPath
					};
				}
				this.resetForUpdateUsingObjectData();
				this.m_objectPathInfo.ObjectPathType=6;
				this.m_objectPathInfo.Name=referenceId;
				delete this.m_objectPathInfo.ParentObjectPathId;
				this.m_parentObjectPath=null;
				return;
			}
			var collectionPropertyPath=clientObject[OfficeExtension.Constants.collectionPropertyPath];
			if (!OfficeExtension.Utility.isNullOrEmptyString(collectionPropertyPath)) {
				var id=OfficeExtension.Utility.tryGetObjectIdFromLoadOrRetrieveResult(value);
				if (!OfficeExtension.Utility.isNullOrUndefined(id)) {
					var propNames=collectionPropertyPath.split(".");
					var parent_1=clientObject.context[propNames[0]];
					for (var i=1; i < propNames.length; i++) {
						parent_1=parent_1[propNames[i]];
					}
					this.resetForUpdateUsingObjectData();
					this.m_parentObjectPath=parent_1._objectPath;
					this.m_objectPathInfo.ParentObjectPathId=this.m_parentObjectPath.objectPathInfo.Id;
					this.m_objectPathInfo.ObjectPathType=5;
					this.m_objectPathInfo.Name="";
					this.m_objectPathInfo.ArgumentInfo.Arguments=[id];
					return;
				}
			}
			var parentIsCollection=this.parentObjectPath && this.parentObjectPath.isCollection;
			var getByIdMethodName=this.getByIdMethodName;
			if (parentIsCollection || !OfficeExtension.Utility.isNullOrEmptyString(getByIdMethodName)) {
				var id=OfficeExtension.Utility.tryGetObjectIdFromLoadOrRetrieveResult(value);
				if (!OfficeExtension.Utility.isNullOrUndefined(id)) {
					this.resetForUpdateUsingObjectData();
					if (!OfficeExtension.Utility.isNullOrEmptyString(getByIdMethodName)) {
						this.m_objectPathInfo.ObjectPathType=3;
						this.m_objectPathInfo.Name=getByIdMethodName;
						this.m_getByIdMethodName=null;
					}
					else {
						this.m_objectPathInfo.ObjectPathType=5;
						this.m_objectPathInfo.Name="";
					}
					this.m_objectPathInfo.ArgumentInfo.Arguments=[id];
					return;
				}
			}
		};
		ObjectPath.prototype.resetForUpdateUsingObjectData=function () {
			this.m_isInvalidAfterRequest=false;
			this.m_isValid=true;
			this.m_isWriteOperation=false;
			this.m_objectPathInfo.ArgumentInfo={};
			this.m_argumentObjectPaths=null;
		};
		ObjectPath.isRestorableObjectPath=function (objectPathType) {
			return (objectPathType===1 ||
				objectPathType===5 ||
				objectPathType===3 ||
				objectPathType===4);
		};
		ObjectPath.copyObjectPathInfo=function (src, dest) {
			dest.Id=src.Id;
			dest.ArgumentInfo=src.ArgumentInfo;
			dest.Name=src.Name;
			dest.ObjectPathType=src.ObjectPathType;
			dest.ParentObjectPathId=src.ParentObjectPathId;
		};
		return ObjectPath;
	}());
	OfficeExtension.ObjectPath=ObjectPath;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ObjectPathFactory=(function () {
		function ObjectPathFactory() {
		}
		ObjectPathFactory.createGlobalObjectObjectPath=function (context) {
			var objectPathInfo={ Id: context._nextId(), ObjectPathType: 1, Name: "" };
			return new OfficeExtension.ObjectPath(objectPathInfo, null, false, false);
		};
		ObjectPathFactory.createNewObjectObjectPath=function (context, typeName, isCollection, isRestrictedResourceAccess) {
			var objectPathInfo={ Id: context._nextId(), ObjectPathType: 2, Name: typeName };
			var ret=new OfficeExtension.ObjectPath(objectPathInfo, null, isCollection, false);
			ret.isRestrictedResourceAccess=isRestrictedResourceAccess;
			return ret;
		};
		ObjectPathFactory.createPropertyObjectPath=function (context, parent, propertyName, isCollection, isInvalidAfterRequest, isRestrictedResourceAccess) {
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 4,
				Name: propertyName,
				ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
			};
			var ret=new OfficeExtension.ObjectPath(objectPathInfo, parent._objectPath, isCollection, isInvalidAfterRequest);
			ret.isRestrictedResourceAccess=isRestrictedResourceAccess;
			return ret;
		};
		ObjectPathFactory.createIndexerObjectPath=function (context, parent, args) {
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 5,
				Name: "",
				ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			objectPathInfo.ArgumentInfo.Arguments=args;
			return new OfficeExtension.ObjectPath(objectPathInfo, parent._objectPath, false, false);
		};
		ObjectPathFactory.createIndexerObjectPathUsingParentPath=function (context, parentObjectPath, args) {
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 5,
				Name: "",
				ParentObjectPathId: parentObjectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			objectPathInfo.ArgumentInfo.Arguments=args;
			return new OfficeExtension.ObjectPath(objectPathInfo, parentObjectPath, false, false);
		};
		ObjectPathFactory.createMethodObjectPath=function (context, parent, methodName, operationType, args, isCollection, isInvalidAfterRequest, getByIdMethodName, isRestrictedResourceAccess) {
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 3,
				Name: methodName,
				ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			var argumentObjectPaths=OfficeExtension.Utility.setMethodArguments(context, objectPathInfo.ArgumentInfo, args);
			var ret=new OfficeExtension.ObjectPath(objectPathInfo, parent._objectPath, isCollection, isInvalidAfterRequest);
			ret.argumentObjectPaths=argumentObjectPaths;
			ret.isWriteOperation=(operationType !=1);
			ret.getByIdMethodName=getByIdMethodName;
			ret.isRestrictedResourceAccess=isRestrictedResourceAccess;
			return ret;
		};
		ObjectPathFactory.createReferenceIdObjectPath=function (context, referenceId) {
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 6,
				Name: referenceId,
				ArgumentInfo: {}
			};
			var ret=new OfficeExtension.ObjectPath(objectPathInfo, null, false, false);
			return ret;
		};
		ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt=function (hasIndexerMethod, context, parent, childItem, index) {
			var id=OfficeExtension.Utility.tryGetObjectIdFromLoadOrRetrieveResult(childItem);
			if (hasIndexerMethod && !OfficeExtension.Utility.isNullOrUndefined(id)) {
				return ObjectPathFactory.createChildItemObjectPathUsingIndexer(context, parent, childItem);
			}
			else {
				return ObjectPathFactory.createChildItemObjectPathUsingGetItemAt(context, parent, childItem, index);
			}
		};
		ObjectPathFactory.createChildItemObjectPathUsingIndexer=function (context, parent, childItem) {
			var id=OfficeExtension.Utility.tryGetObjectIdFromLoadOrRetrieveResult(childItem);
			var objectPathInfo=objectPathInfo=				{
					Id: context._nextId(),
					ObjectPathType: 5,
					Name: "",
					ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
					ArgumentInfo: {}
				};
			objectPathInfo.ArgumentInfo.Arguments=[id];
			return new OfficeExtension.ObjectPath(objectPathInfo, parent._objectPath, false, false);
		};
		ObjectPathFactory.createChildItemObjectPathUsingGetItemAt=function (context, parent, childItem, index) {
			var indexFromServer=childItem[OfficeExtension.Constants.index];
			if (indexFromServer) {
				index=indexFromServer;
			}
			var objectPathInfo={
				Id: context._nextId(),
				ObjectPathType: 3,
				Name: OfficeExtension.Constants.getItemAt,
				ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
				ArgumentInfo: {}
			};
			objectPathInfo.ArgumentInfo.Arguments=[index];
			return new OfficeExtension.ObjectPath(objectPathInfo, parent._objectPath, false, false);
		};
		return ObjectPathFactory;
	}());
	OfficeExtension.ObjectPathFactory=ObjectPathFactory;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var OfficeJsRequestExecutor=(function () {
		function OfficeJsRequestExecutor(context) {
			this.m_context=context;
		}
		OfficeJsRequestExecutor.prototype.executeAsync=function (customData, requestFlags, requestMessage) {
			var _this=this;
			var messageSafearray=OfficeExtension.RichApiMessageUtility.buildMessageArrayForIRequestExecutor(customData, requestFlags, requestMessage, OfficeJsRequestExecutor.SourceLibHeaderValue);
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				OSF.DDA.RichApi.executeRichApiRequestAsync(messageSafearray, function (result) {
					OfficeExtension.Utility.log("Response:");
					OfficeExtension.Utility.log(JSON.stringify(result));
					var response;
					if (result.status=="succeeded") {
						response=OfficeExtension.RichApiMessageUtility.buildResponseOnSuccess(OfficeExtension.RichApiMessageUtility.getResponseBody(result), OfficeExtension.RichApiMessageUtility.getResponseHeaders(result));
					}
					else {
						response=OfficeExtension.RichApiMessageUtility.buildResponseOnError(result.error.code, result.error.message);
						_this.m_context._processOfficeJsErrorResponse(result.error.code, response);
					}
					resolve(response);
				});
			});
		};
		OfficeJsRequestExecutor.SourceLibHeaderValue="officejs";
		return OfficeJsRequestExecutor;
	}());
	OfficeExtension.OfficeJsRequestExecutor=OfficeJsRequestExecutor;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var _Internal;
	(function (_Internal) {
		_Internal.OfficeRequire=function () {
			return null;
		}();
	})(_Internal=OfficeExtension._Internal || (OfficeExtension._Internal={}));
	var _Internal;
	(function (_Internal) {
		var PromiseImpl;
		(function (PromiseImpl) {
			function Init() {
				return (function () {
					"use strict";
					function lib$es6$promise$utils$$objectOrFunction(x) {
						return typeof x==='function' || (typeof x==='object' && x !==null);
					}
					function lib$es6$promise$utils$$isFunction(x) {
						return typeof x==='function';
					}
					function lib$es6$promise$utils$$isMaybeThenable(x) {
						return typeof x==='object' && x !==null;
					}
					var lib$es6$promise$utils$$_isArray;
					if (!Array.isArray) {
						lib$es6$promise$utils$$_isArray=function (x) {
							return Object.prototype.toString.call(x)==='[object Array]';
						};
					}
					else {
						lib$es6$promise$utils$$_isArray=Array.isArray;
					}
					var lib$es6$promise$utils$$isArray=lib$es6$promise$utils$$_isArray;
					var lib$es6$promise$asap$$len=0;
					var lib$es6$promise$asap$$toString={}.toString;
					var lib$es6$promise$asap$$vertxNext;
					var lib$es6$promise$asap$$customSchedulerFn;
					var lib$es6$promise$asap$$asap=function asap(callback, arg) {
						lib$es6$promise$asap$$queue[lib$es6$promise$asap$$len]=callback;
						lib$es6$promise$asap$$queue[lib$es6$promise$asap$$len+1]=arg;
						lib$es6$promise$asap$$len+=2;
						if (lib$es6$promise$asap$$len===2) {
							if (lib$es6$promise$asap$$customSchedulerFn) {
								lib$es6$promise$asap$$customSchedulerFn(lib$es6$promise$asap$$flush);
							}
							else {
								lib$es6$promise$asap$$scheduleFlush();
							}
						}
					};
					function lib$es6$promise$asap$$setScheduler(scheduleFn) {
						lib$es6$promise$asap$$customSchedulerFn=scheduleFn;
					}
					function lib$es6$promise$asap$$setAsap(asapFn) {
						lib$es6$promise$asap$$asap=asapFn;
					}
					var lib$es6$promise$asap$$browserWindow=(typeof window !=='undefined') ? window : undefined;
					var lib$es6$promise$asap$$browserGlobal=lib$es6$promise$asap$$browserWindow || {};
					var lib$es6$promise$asap$$BrowserMutationObserver=lib$es6$promise$asap$$browserGlobal.MutationObserver || lib$es6$promise$asap$$browserGlobal.WebKitMutationObserver;
					var lib$es6$promise$asap$$isNode=typeof process !=='undefined' && {}.toString.call(process)==='[object process]';
					var lib$es6$promise$asap$$isWorker=typeof Uint8ClampedArray !=='undefined' &&
						typeof importScripts !=='undefined' &&
						typeof MessageChannel !=='undefined';
					function lib$es6$promise$asap$$useNextTick() {
						var nextTick=process.nextTick;
						var version=process.versions.node.match(/^(?:(\d+)\.)?(?:(\d+)\.)?(\*|\d+)$/);
						if (Array.isArray(version) && version[1]==='0' && version[2]==='10') {
							nextTick=setImmediate;
						}
						return function () {
							nextTick(lib$es6$promise$asap$$flush);
						};
					}
					function lib$es6$promise$asap$$useVertxTimer() {
						return function () {
							lib$es6$promise$asap$$vertxNext(lib$es6$promise$asap$$flush);
						};
					}
					function lib$es6$promise$asap$$useMutationObserver() {
						var iterations=0;
						var observer=new lib$es6$promise$asap$$BrowserMutationObserver(lib$es6$promise$asap$$flush);
						var node=document.createTextNode('');
						observer.observe(node, { characterData: true });
						return function () {
							node.data=(iterations=++iterations % 2);
						};
					}
					function lib$es6$promise$asap$$useMessageChannel() {
						var channel=new MessageChannel();
						channel.port1.onmessage=lib$es6$promise$asap$$flush;
						return function () {
							channel.port2.postMessage(0);
						};
					}
					function lib$es6$promise$asap$$useSetTimeout() {
						return function () {
							setTimeout(lib$es6$promise$asap$$flush, 1);
						};
					}
					var lib$es6$promise$asap$$queue=new Array(1000);
					function lib$es6$promise$asap$$flush() {
						for (var i=0; i < lib$es6$promise$asap$$len; i+=2) {
							var callback=lib$es6$promise$asap$$queue[i];
							var arg=lib$es6$promise$asap$$queue[i+1];
							callback(arg);
							lib$es6$promise$asap$$queue[i]=undefined;
							lib$es6$promise$asap$$queue[i+1]=undefined;
						}
						lib$es6$promise$asap$$len=0;
					}
					var lib$es6$promise$asap$$scheduleFlush;
					if (lib$es6$promise$asap$$isNode) {
						lib$es6$promise$asap$$scheduleFlush=lib$es6$promise$asap$$useNextTick();
					}
					else if (lib$es6$promise$asap$$BrowserMutationObserver) {
						lib$es6$promise$asap$$scheduleFlush=lib$es6$promise$asap$$useMutationObserver();
					}
					else if (lib$es6$promise$asap$$isWorker) {
						lib$es6$promise$asap$$scheduleFlush=lib$es6$promise$asap$$useMessageChannel();
					}
					else {
						lib$es6$promise$asap$$scheduleFlush=lib$es6$promise$asap$$useSetTimeout();
					}
					function lib$es6$promise$$internal$$noop() { }
					var lib$es6$promise$$internal$$PENDING=void 0;
					var lib$es6$promise$$internal$$FULFILLED=1;
					var lib$es6$promise$$internal$$REJECTED=2;
					var lib$es6$promise$$internal$$GET_THEN_ERROR=new lib$es6$promise$$internal$$ErrorObject();
					function lib$es6$promise$$internal$$selfFullfillment() {
						return new TypeError("You cannot resolve a promise with itself");
					}
					function lib$es6$promise$$internal$$cannotReturnOwn() {
						return new TypeError('A promises callback cannot return that same promise.');
					}
					function lib$es6$promise$$internal$$getThen(promise) {
						try {
							return promise.then;
						}
						catch (error) {
							lib$es6$promise$$internal$$GET_THEN_ERROR.error=error;
							return lib$es6$promise$$internal$$GET_THEN_ERROR;
						}
					}
					function lib$es6$promise$$internal$$tryThen(then, value, fulfillmentHandler, rejectionHandler) {
						try {
							then.call(value, fulfillmentHandler, rejectionHandler);
						}
						catch (e) {
							return e;
						}
					}
					function lib$es6$promise$$internal$$handleForeignThenable(promise, thenable, then) {
						lib$es6$promise$asap$$asap(function (promise) {
							var sealed=false;
							var error=lib$es6$promise$$internal$$tryThen(then, thenable, function (value) {
								if (sealed) {
									return;
								}
								sealed=true;
								if (thenable !==value) {
									lib$es6$promise$$internal$$resolve(promise, value);
								}
								else {
									lib$es6$promise$$internal$$fulfill(promise, value);
								}
							}, function (reason) {
								if (sealed) {
									return;
								}
								sealed=true;
								lib$es6$promise$$internal$$reject(promise, reason);
							}, 'Settle: '+(promise._label || ' unknown promise'));
							if (!sealed && error) {
								sealed=true;
								lib$es6$promise$$internal$$reject(promise, error);
							}
						}, promise);
					}
					function lib$es6$promise$$internal$$handleOwnThenable(promise, thenable) {
						if (thenable._state===lib$es6$promise$$internal$$FULFILLED) {
							lib$es6$promise$$internal$$fulfill(promise, thenable._result);
						}
						else if (thenable._state===lib$es6$promise$$internal$$REJECTED) {
							lib$es6$promise$$internal$$reject(promise, thenable._result);
						}
						else {
							lib$es6$promise$$internal$$subscribe(thenable, undefined, function (value) {
								lib$es6$promise$$internal$$resolve(promise, value);
							}, function (reason) {
								lib$es6$promise$$internal$$reject(promise, reason);
							});
						}
					}
					function lib$es6$promise$$internal$$handleMaybeThenable(promise, maybeThenable) {
						if (maybeThenable.constructor===promise.constructor) {
							lib$es6$promise$$internal$$handleOwnThenable(promise, maybeThenable);
						}
						else {
							var then=lib$es6$promise$$internal$$getThen(maybeThenable);
							if (then===lib$es6$promise$$internal$$GET_THEN_ERROR) {
								lib$es6$promise$$internal$$reject(promise, lib$es6$promise$$internal$$GET_THEN_ERROR.error);
							}
							else if (then===undefined) {
								lib$es6$promise$$internal$$fulfill(promise, maybeThenable);
							}
							else if (lib$es6$promise$utils$$isFunction(then)) {
								lib$es6$promise$$internal$$handleForeignThenable(promise, maybeThenable, then);
							}
							else {
								lib$es6$promise$$internal$$fulfill(promise, maybeThenable);
							}
						}
					}
					function lib$es6$promise$$internal$$resolve(promise, value) {
						if (promise===value) {
							lib$es6$promise$$internal$$reject(promise, lib$es6$promise$$internal$$selfFullfillment());
						}
						else if (lib$es6$promise$utils$$objectOrFunction(value)) {
							lib$es6$promise$$internal$$handleMaybeThenable(promise, value);
						}
						else {
							lib$es6$promise$$internal$$fulfill(promise, value);
						}
					}
					function lib$es6$promise$$internal$$publishRejection(promise) {
						if (promise._onerror) {
							promise._onerror(promise._result);
						}
						lib$es6$promise$$internal$$publish(promise);
					}
					function lib$es6$promise$$internal$$fulfill(promise, value) {
						if (promise._state !==lib$es6$promise$$internal$$PENDING) {
							return;
						}
						promise._result=value;
						promise._state=lib$es6$promise$$internal$$FULFILLED;
						if (promise._subscribers.length !==0) {
							lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publish, promise);
						}
					}
					function lib$es6$promise$$internal$$reject(promise, reason) {
						if (promise._state !==lib$es6$promise$$internal$$PENDING) {
							return;
						}
						promise._state=lib$es6$promise$$internal$$REJECTED;
						promise._result=reason;
						lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publishRejection, promise);
					}
					function lib$es6$promise$$internal$$subscribe(parent, child, onFulfillment, onRejection) {
						var subscribers=parent._subscribers;
						var length=subscribers.length;
						parent._onerror=null;
						subscribers[length]=child;
						subscribers[length+lib$es6$promise$$internal$$FULFILLED]=onFulfillment;
						subscribers[length+lib$es6$promise$$internal$$REJECTED]=onRejection;
						if (length===0 && parent._state) {
							lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publish, parent);
						}
					}
					function lib$es6$promise$$internal$$publish(promise) {
						var subscribers=promise._subscribers;
						var settled=promise._state;
						if (subscribers.length===0) {
							return;
						}
						var child, callback, detail=promise._result;
						for (var i=0; i < subscribers.length; i+=3) {
							child=subscribers[i];
							callback=subscribers[i+settled];
							if (child) {
								lib$es6$promise$$internal$$invokeCallback(settled, child, callback, detail);
							}
							else {
								callback(detail);
							}
						}
						promise._subscribers.length=0;
					}
					function lib$es6$promise$$internal$$ErrorObject() {
						this.error=null;
					}
					var lib$es6$promise$$internal$$TRY_CATCH_ERROR=new lib$es6$promise$$internal$$ErrorObject();
					function lib$es6$promise$$internal$$tryCatch(callback, detail) {
						try {
							return callback(detail);
						}
						catch (e) {
							lib$es6$promise$$internal$$TRY_CATCH_ERROR.error=e;
							return lib$es6$promise$$internal$$TRY_CATCH_ERROR;
						}
					}
					function lib$es6$promise$$internal$$invokeCallback(settled, promise, callback, detail) {
						var hasCallback=lib$es6$promise$utils$$isFunction(callback), value, error, succeeded, failed;
						if (hasCallback) {
							value=lib$es6$promise$$internal$$tryCatch(callback, detail);
							if (value===lib$es6$promise$$internal$$TRY_CATCH_ERROR) {
								failed=true;
								error=value.error;
								value=null;
							}
							else {
								succeeded=true;
							}
							if (promise===value) {
								lib$es6$promise$$internal$$reject(promise, lib$es6$promise$$internal$$cannotReturnOwn());
								return;
							}
						}
						else {
							value=detail;
							succeeded=true;
						}
						if (promise._state !==lib$es6$promise$$internal$$PENDING) {
						}
						else if (hasCallback && succeeded) {
							lib$es6$promise$$internal$$resolve(promise, value);
						}
						else if (failed) {
							lib$es6$promise$$internal$$reject(promise, error);
						}
						else if (settled===lib$es6$promise$$internal$$FULFILLED) {
							lib$es6$promise$$internal$$fulfill(promise, value);
						}
						else if (settled===lib$es6$promise$$internal$$REJECTED) {
							lib$es6$promise$$internal$$reject(promise, value);
						}
					}
					function lib$es6$promise$$internal$$initializePromise(promise, resolver) {
						try {
							resolver(function resolvePromise(value) {
								lib$es6$promise$$internal$$resolve(promise, value);
							}, function rejectPromise(reason) {
								lib$es6$promise$$internal$$reject(promise, reason);
							});
						}
						catch (e) {
							lib$es6$promise$$internal$$reject(promise, e);
						}
					}
					function lib$es6$promise$enumerator$$Enumerator(Constructor, input) {
						var enumerator=this;
						enumerator._instanceConstructor=Constructor;
						enumerator.promise=new Constructor(lib$es6$promise$$internal$$noop);
						if (enumerator._validateInput(input)) {
							enumerator._input=input;
							enumerator.length=input.length;
							enumerator._remaining=input.length;
							enumerator._init();
							if (enumerator.length===0) {
								lib$es6$promise$$internal$$fulfill(enumerator.promise, enumerator._result);
							}
							else {
								enumerator.length=enumerator.length || 0;
								enumerator._enumerate();
								if (enumerator._remaining===0) {
									lib$es6$promise$$internal$$fulfill(enumerator.promise, enumerator._result);
								}
							}
						}
						else {
							lib$es6$promise$$internal$$reject(enumerator.promise, enumerator._validationError());
						}
					}
					lib$es6$promise$enumerator$$Enumerator.prototype._validateInput=function (input) {
						return lib$es6$promise$utils$$isArray(input);
					};
					lib$es6$promise$enumerator$$Enumerator.prototype._validationError=function () {
						return new _Internal.Error('Array Methods must be provided an Array');
					};
					lib$es6$promise$enumerator$$Enumerator.prototype._init=function () {
						this._result=new Array(this.length);
					};
					var lib$es6$promise$enumerator$$default=lib$es6$promise$enumerator$$Enumerator;
					lib$es6$promise$enumerator$$Enumerator.prototype._enumerate=function () {
						var enumerator=this;
						var length=enumerator.length;
						var promise=enumerator.promise;
						var input=enumerator._input;
						for (var i=0; promise._state===lib$es6$promise$$internal$$PENDING && i < length; i++) {
							enumerator._eachEntry(input[i], i);
						}
					};
					lib$es6$promise$enumerator$$Enumerator.prototype._eachEntry=function (entry, i) {
						var enumerator=this;
						var c=enumerator._instanceConstructor;
						if (lib$es6$promise$utils$$isMaybeThenable(entry)) {
							if (entry.constructor===c && entry._state !==lib$es6$promise$$internal$$PENDING) {
								entry._onerror=null;
								enumerator._settledAt(entry._state, i, entry._result);
							}
							else {
								enumerator._willSettleAt(c.resolve(entry), i);
							}
						}
						else {
							enumerator._remaining--;
							enumerator._result[i]=entry;
						}
					};
					lib$es6$promise$enumerator$$Enumerator.prototype._settledAt=function (state, i, value) {
						var enumerator=this;
						var promise=enumerator.promise;
						if (promise._state===lib$es6$promise$$internal$$PENDING) {
							enumerator._remaining--;
							if (state===lib$es6$promise$$internal$$REJECTED) {
								lib$es6$promise$$internal$$reject(promise, value);
							}
							else {
								enumerator._result[i]=value;
							}
						}
						if (enumerator._remaining===0) {
							lib$es6$promise$$internal$$fulfill(promise, enumerator._result);
						}
					};
					lib$es6$promise$enumerator$$Enumerator.prototype._willSettleAt=function (promise, i) {
						var enumerator=this;
						lib$es6$promise$$internal$$subscribe(promise, undefined, function (value) {
							enumerator._settledAt(lib$es6$promise$$internal$$FULFILLED, i, value);
						}, function (reason) {
							enumerator._settledAt(lib$es6$promise$$internal$$REJECTED, i, reason);
						});
					};
					function lib$es6$promise$promise$all$$all(entries) {
						return new lib$es6$promise$enumerator$$default(this, entries).promise;
					}
					var lib$es6$promise$promise$all$$default=lib$es6$promise$promise$all$$all;
					function lib$es6$promise$promise$race$$race(entries) {
						var Constructor=this;
						var promise=new Constructor(lib$es6$promise$$internal$$noop);
						if (!lib$es6$promise$utils$$isArray(entries)) {
							lib$es6$promise$$internal$$reject(promise, new TypeError('You must pass an array to race.'));
							return promise;
						}
						var length=entries.length;
						function onFulfillment(value) {
							lib$es6$promise$$internal$$resolve(promise, value);
						}
						function onRejection(reason) {
							lib$es6$promise$$internal$$reject(promise, reason);
						}
						for (var i=0; promise._state===lib$es6$promise$$internal$$PENDING && i < length; i++) {
							lib$es6$promise$$internal$$subscribe(Constructor.resolve(entries[i]), undefined, onFulfillment, onRejection);
						}
						return promise;
					}
					var lib$es6$promise$promise$race$$default=lib$es6$promise$promise$race$$race;
					function lib$es6$promise$promise$resolve$$resolve(object) {
						var Constructor=this;
						if (object && typeof object==='object' && object.constructor===Constructor) {
							return object;
						}
						var promise=new Constructor(lib$es6$promise$$internal$$noop);
						lib$es6$promise$$internal$$resolve(promise, object);
						return promise;
					}
					var lib$es6$promise$promise$resolve$$default=lib$es6$promise$promise$resolve$$resolve;
					function lib$es6$promise$promise$reject$$reject(reason) {
						var Constructor=this;
						var promise=new Constructor(lib$es6$promise$$internal$$noop);
						lib$es6$promise$$internal$$reject(promise, reason);
						return promise;
					}
					var lib$es6$promise$promise$reject$$default=lib$es6$promise$promise$reject$$reject;
					var lib$es6$promise$promise$$counter=0;
					function lib$es6$promise$promise$$needsResolver() {
						throw new TypeError('You must pass a resolver function as the first argument to the promise constructor');
					}
					function lib$es6$promise$promise$$needsNew() {
						throw new TypeError("Failed to construct 'Promise': Please use the 'new' operator, this object constructor cannot be called as a function.");
					}
					var lib$es6$promise$promise$$default=lib$es6$promise$promise$$Promise;
					function lib$es6$promise$promise$$Promise(resolver) {
						this._id=lib$es6$promise$promise$$counter++;
						this._state=undefined;
						this._result=undefined;
						this._subscribers=[];
						if (lib$es6$promise$$internal$$noop !==resolver) {
							if (!lib$es6$promise$utils$$isFunction(resolver)) {
								lib$es6$promise$promise$$needsResolver();
							}
							if (!(this instanceof lib$es6$promise$promise$$Promise)) {
								lib$es6$promise$promise$$needsNew();
							}
							lib$es6$promise$$internal$$initializePromise(this, resolver);
						}
					}
					lib$es6$promise$promise$$Promise.all=lib$es6$promise$promise$all$$default;
					lib$es6$promise$promise$$Promise.race=lib$es6$promise$promise$race$$default;
					lib$es6$promise$promise$$Promise.resolve=lib$es6$promise$promise$resolve$$default;
					lib$es6$promise$promise$$Promise.reject=lib$es6$promise$promise$reject$$default;
					lib$es6$promise$promise$$Promise._setScheduler=lib$es6$promise$asap$$setScheduler;
					lib$es6$promise$promise$$Promise._setAsap=lib$es6$promise$asap$$setAsap;
					lib$es6$promise$promise$$Promise._asap=lib$es6$promise$asap$$asap;
					lib$es6$promise$promise$$Promise.prototype={
						constructor: lib$es6$promise$promise$$Promise,
						then: function (onFulfillment, onRejection) {
							var parent=this;
							var state=parent._state;
							if (state===lib$es6$promise$$internal$$FULFILLED && !onFulfillment || state===lib$es6$promise$$internal$$REJECTED && !onRejection) {
								return this;
							}
							var child=new this.constructor(lib$es6$promise$$internal$$noop);
							var result=parent._result;
							if (state) {
								var callback=arguments[state - 1];
								lib$es6$promise$asap$$asap(function () {
									lib$es6$promise$$internal$$invokeCallback(state, child, callback, result);
								});
							}
							else {
								lib$es6$promise$$internal$$subscribe(parent, child, onFulfillment, onRejection);
							}
							return child;
						},
						'catch': function (onRejection) {
							return this.then(null, onRejection);
						}
					};
					return lib$es6$promise$promise$$default;
				}).call(this);
			}
			PromiseImpl.Init=Init;
		})(PromiseImpl=_Internal.PromiseImpl || (_Internal.PromiseImpl={}));
	})(_Internal=OfficeExtension._Internal || (OfficeExtension._Internal={}));
	var _Internal;
	(function (_Internal) {
		function isEdgeLessThan14() {
			var userAgent=window.navigator.userAgent;
			var versionIdx=userAgent.indexOf("Edge/");
			if (versionIdx >=0) {
				userAgent=userAgent.substring(versionIdx+5, userAgent.length);
				if (userAgent < "14.14393")
					return true;
				else
					return false;
			}
			return false;
		}
		function determinePromise() {
			if (typeof (window)==="undefined" && typeof (Promise)==="function") {
				return Promise;
			}
			if (typeof (window) !=="undefined" && window.Promise) {
				if (isEdgeLessThan14()) {
					return _Internal.PromiseImpl.Init();
				}
				else {
					return window.Promise;
				}
			}
			else {
				return _Internal.PromiseImpl.Init();
			}
		}
		_Internal.OfficePromise=determinePromise();
	})(_Internal=OfficeExtension._Internal || (OfficeExtension._Internal={}));
	var OfficePromise=_Internal.OfficePromise;
	OfficeExtension.Promise=OfficePromise;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var TrackedObjects=(function () {
		function TrackedObjects(context) {
			this._autoCleanupList={};
			this.m_context=context;
		}
		TrackedObjects.prototype.add=function (param) {
			var _this=this;
			if (Array.isArray(param)) {
				param.forEach(function (item) { return _this._addCommon(item, true); });
			}
			else {
				this._addCommon(param, true);
			}
		};
		TrackedObjects.prototype._autoAdd=function (object) {
			this._addCommon(object, false);
			this._autoCleanupList[object._objectPath.objectPathInfo.Id]=object;
		};
		TrackedObjects.prototype._autoTrackIfNecessaryWhenHandleObjectResultValue=function (object, resultValue) {
			var shouldAutoTrack=(this.m_context._autoCleanup &&
				!object[OfficeExtension.Constants.isTracked] &&
				object !==this.m_context._rootObject &&
				resultValue &&
				!OfficeExtension.Utility.isNullOrEmptyString(resultValue[OfficeExtension.Constants.referenceId]));
			if (shouldAutoTrack) {
				this._autoCleanupList[object._objectPath.objectPathInfo.Id]=object;
				object[OfficeExtension.Constants.isTracked]=true;
			}
		};
		TrackedObjects.prototype._addCommon=function (object, isExplicitlyAdded) {
			if (object[OfficeExtension.Constants.isTracked]) {
				if (isExplicitlyAdded && this.m_context._autoCleanup) {
					delete this._autoCleanupList[object._objectPath.objectPathInfo.Id];
				}
				return;
			}
			var referenceId=object[OfficeExtension.Constants.referenceId];
			if (OfficeExtension.Utility.isNullOrEmptyString(referenceId) && object._KeepReference) {
				object._KeepReference();
				OfficeExtension.ActionFactory.createInstantiateAction(this.m_context, object);
				if (isExplicitlyAdded && this.m_context._autoCleanup) {
					delete this._autoCleanupList[object._objectPath.objectPathInfo.Id];
				}
				object[OfficeExtension.Constants.isTracked]=true;
			}
		};
		TrackedObjects.prototype.remove=function (param) {
			var _this=this;
			if (Array.isArray(param)) {
				param.forEach(function (item) { return _this._removeCommon(item); });
			}
			else {
				this._removeCommon(param);
			}
		};
		TrackedObjects.prototype._removeCommon=function (object) {
			var referenceId=object[OfficeExtension.Constants.referenceId];
			if (!OfficeExtension.Utility.isNullOrEmptyString(referenceId)) {
				var rootObject=this.m_context._rootObject;
				if (rootObject._RemoveReference) {
					rootObject._RemoveReference(referenceId);
				}
				delete object[OfficeExtension.Constants.isTracked];
			}
		};
		TrackedObjects.prototype._retrieveAndClearAutoCleanupList=function () {
			var list=this._autoCleanupList;
			this._autoCleanupList={};
			return list;
		};
		return TrackedObjects;
	}());
	OfficeExtension.TrackedObjects=TrackedObjects;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var RequestPrettyPrinter=(function () {
		function RequestPrettyPrinter(globalObjName, referencedObjectPaths, actions, showDispose) {
			if (!globalObjName) {
				globalObjName="root";
			}
			this.m_globalObjName=globalObjName;
			this.m_referencedObjectPaths=referencedObjectPaths;
			this.m_actions=actions;
			this.m_statements=[];
			this.m_variableNameForObjectPathMap={};
			this.m_variableNameToObjectPathMap={};
			this.m_declaredObjectPathMap={};
			this.m_showDispose=showDispose;
		}
		RequestPrettyPrinter.prototype.process=function () {
			if (this.m_showDispose) {
				OfficeExtension.ClientRequest._calculateLastUsedObjectPathIds(this.m_actions);
			}
			for (var i=0; i < this.m_actions.length; i++) {
				this.processOneAction(this.m_actions[i]);
			}
			return this.m_statements;
		};
		RequestPrettyPrinter.prototype.processForDebugStatementInfo=function (actionIndex) {
			if (this.m_showDispose) {
				OfficeExtension.ClientRequest._calculateLastUsedObjectPathIds(this.m_actions);
			}
			var surroundingCount=5;
			this.m_statements=[];
			var oneStatement="";
			var statementIndex=-1;
			for (var i=0; i < this.m_actions.length; i++) {
				this.processOneAction(this.m_actions[i]);
				if (actionIndex==i) {
					statementIndex=this.m_statements.length - 1;
				}
				if (statementIndex >=0 && this.m_statements.length > statementIndex+surroundingCount+1) {
					break;
				}
			}
			if (statementIndex < 0) {
				return null;
			}
			var startIndex=statementIndex - surroundingCount;
			if (startIndex < 0) {
				startIndex=0;
			}
			var endIndex=statementIndex+1+surroundingCount;
			if (endIndex > this.m_statements.length) {
				endIndex=this.m_statements.length;
			}
			var surroundingStatements=[];
			if (startIndex !=0) {
				surroundingStatements.push("...");
			}
			for (var i_1=startIndex; i_1 < statementIndex; i_1++) {
				surroundingStatements.push(this.m_statements[i_1]);
			}
			surroundingStatements.push("// >>>>>");
			surroundingStatements.push(this.m_statements[statementIndex]);
			surroundingStatements.push("// <<<<<");
			for (var i_2=statementIndex+1; i_2 < endIndex; i_2++) {
				surroundingStatements.push(this.m_statements[i_2]);
			}
			if (endIndex < this.m_statements.length) {
				surroundingStatements.push("...");
			}
			return {
				statement: this.m_statements[statementIndex],
				surroundingStatements: surroundingStatements
			};
		};
		RequestPrettyPrinter.prototype.processOneAction=function (action) {
			var actionInfo=action.actionInfo;
			switch (actionInfo.ActionType) {
				case 1:
					this.processInstantiateAction(action);
					break;
				case 3:
					this.processMethodAction(action);
					break;
				case 2:
					this.processQueryAction(action);
					break;
				case 7:
					this.processQueryAsJsonAction(action);
					break;
				case 6:
					this.processRecursiveQueryAction(action);
					break;
				case 4:
					this.processSetPropertyAction(action);
					break;
				case 5:
					this.processTraceAction(action);
					break;
				case 8:
					this.processEnsureUnchangedAction(action);
					break;
				case 9:
					this.processUpdateAction(action);
					break;
			}
		};
		RequestPrettyPrinter.prototype.processInstantiateAction=function (action) {
			var objId=action.actionInfo.ObjectPathId;
			var objPath=this.m_referencedObjectPaths[objId];
			var varName=this.getObjVarName(objId);
			if (!this.m_declaredObjectPathMap[objId]) {
				var statement="var "+varName+"="+this.buildObjectPathExpressionWithParent(objPath)+";";
				statement=this.appendDisposeCommentIfRelevant(statement, action);
				this.m_statements.push(statement);
				this.m_declaredObjectPathMap[objId]=varName;
			}
			else {
				var statement="// Instantiate {"+varName+"}";
				statement=this.appendDisposeCommentIfRelevant(statement, action);
				this.m_statements.push(statement);
			}
		};
		RequestPrettyPrinter.prototype.processMethodAction=function (action) {
			var methodName=action.actionInfo.Name;
			if (methodName==="_KeepReference") {
				methodName="track";
			}
			var statement=this.getObjVarName(action.actionInfo.ObjectPathId)+"."+OfficeExtension.Utility._toCamelLowerCase(methodName)+"("+this.buildArgumentsExpression(action.actionInfo.ArgumentInfo)+");";
			statement=this.appendDisposeCommentIfRelevant(statement, action);
			this.m_statements.push(statement);
		};
		RequestPrettyPrinter.prototype.processQueryAction=function (action) {
			var queryExp=this.buildQueryExpression(action);
			var statement=this.getObjVarName(action.actionInfo.ObjectPathId)+".load("+queryExp+");";
			statement=this.appendDisposeCommentIfRelevant(statement, action);
			this.m_statements.push(statement);
		};
		RequestPrettyPrinter.prototype.processQueryAsJsonAction=function (action) {
			var queryExp=this.buildQueryExpression(action);
			var statement=this.getObjVarName(action.actionInfo.ObjectPathId)+".retrieve("+queryExp+");";
			statement=this.appendDisposeCommentIfRelevant(statement, action);
			this.m_statements.push(statement);
		};
		RequestPrettyPrinter.prototype.processRecursiveQueryAction=function (action) {
			var queryExp="";
			if (action.actionInfo.RecursiveQueryInfo) {
				queryExp=JSON.stringify(action.actionInfo.RecursiveQueryInfo);
			}
			var statement=this.getObjVarName(action.actionInfo.ObjectPathId)+".loadRecursive("+queryExp+");";
			statement=this.appendDisposeCommentIfRelevant(statement, action);
			this.m_statements.push(statement);
		};
		RequestPrettyPrinter.prototype.processSetPropertyAction=function (action) {
			var statement=this.getObjVarName(action.actionInfo.ObjectPathId)+"."+OfficeExtension.Utility._toCamelLowerCase(action.actionInfo.Name)+"="+this.buildArgumentsExpression(action.actionInfo.ArgumentInfo)+";";
			statement=this.appendDisposeCommentIfRelevant(statement, action);
			this.m_statements.push(statement);
		};
		RequestPrettyPrinter.prototype.processTraceAction=function (action) {
			var statement="context.trace();";
			statement=this.appendDisposeCommentIfRelevant(statement, action);
			this.m_statements.push(statement);
		};
		RequestPrettyPrinter.prototype.processEnsureUnchangedAction=function (action) {
			var statement=this.getObjVarName(action.actionInfo.ObjectPathId)+".ensureUnchanged("+JSON.stringify(action.actionInfo.ObjectState)+");";
			statement=this.appendDisposeCommentIfRelevant(statement, action);
			this.m_statements.push(statement);
		};
		RequestPrettyPrinter.prototype.processUpdateAction=function (action) {
			var statement=this.getObjVarName(action.actionInfo.ObjectPathId)+".update("+JSON.stringify(action.actionInfo.ObjectState)+");";
			statement=this.appendDisposeCommentIfRelevant(statement, action);
			this.m_statements.push(statement);
		};
		RequestPrettyPrinter.prototype.appendDisposeCommentIfRelevant=function (statement, action) {
			var _this=this;
			if (this.m_showDispose) {
				var lastUsedObjectPathIds=action.actionInfo.L;
				if (lastUsedObjectPathIds && lastUsedObjectPathIds.length > 0) {
					var objectNamesToDispose=lastUsedObjectPathIds.map(function (item) { return _this.getObjVarName(item); }).join(", ");
					return statement+" // And then dispose {"+objectNamesToDispose+"}";
				}
			}
			return statement;
		};
		RequestPrettyPrinter.prototype.buildQueryExpression=function (action) {
			if (action.actionInfo.QueryInfo) {
				var option={};
				option.select=action.actionInfo.QueryInfo.Select;
				option.expand=action.actionInfo.QueryInfo.Expand;
				option.skip=action.actionInfo.QueryInfo.Skip;
				option.top=action.actionInfo.QueryInfo.Top;
				if (typeof (option.top)==="undefined" && typeof (option.skip)==="undefined" && typeof (option.expand)==="undefined") {
					if (typeof (option.select)==="undefined") {
						return "";
					}
					else {
						return JSON.stringify(option.select);
					}
				}
				else {
					return JSON.stringify(option);
				}
			}
			return "";
		};
		RequestPrettyPrinter.prototype.buildObjectPathExpressionWithParent=function (objPath) {
			var hasParent=objPath.objectPathInfo.ObjectPathType==5 ||
				objPath.objectPathInfo.ObjectPathType==3 ||
				objPath.objectPathInfo.ObjectPathType==4;
			if (hasParent && objPath.objectPathInfo.ParentObjectPathId) {
				return this.getObjVarName(objPath.objectPathInfo.ParentObjectPathId)+"."+this.buildObjectPathExpression(objPath);
			}
			return this.buildObjectPathExpression(objPath);
		};
		RequestPrettyPrinter.prototype.buildObjectPathExpression=function (objPath) {
			switch (objPath.objectPathInfo.ObjectPathType) {
				case 1:
					return "context."+this.m_globalObjName;
				case 5:
					return "getItem("+this.buildArgumentsExpression(objPath.objectPathInfo.ArgumentInfo)+")";
				case 3:
					return OfficeExtension.Utility._toCamelLowerCase(objPath.objectPathInfo.Name)+"("+this.buildArgumentsExpression(objPath.objectPathInfo.ArgumentInfo)+")";
				case 2:
					return objPath.objectPathInfo.Name+".newObject()";
				case 7:
					return "null";
				case 4:
					return OfficeExtension.Utility._toCamelLowerCase(objPath.objectPathInfo.Name);
				case 6:
					return "context."+this.m_globalObjName+"._getObjectByReferenceId("+JSON.stringify(objPath.objectPathInfo.Name)+")";
			}
		};
		RequestPrettyPrinter.prototype.buildArgumentsExpression=function (args) {
			var ret="";
			if (!args.Arguments) {
				return ret;
			}
			for (var i=0; i < args.Arguments.length; i++) {
				if (i > 0) {
					ret=ret+", ";
				}
				ret=ret+this.buildArgumentLiteral(args.Arguments[i], args.ReferencedObjectPathIds ? args.ReferencedObjectPathIds[i] : null);
			}
			if (ret==="undefined") {
				ret="";
			}
			return ret;
		};
		RequestPrettyPrinter.prototype.buildArgumentLiteral=function (value, objectPathId) {
			if (typeof value=="number" && value===objectPathId) {
				return this.getObjVarName(objectPathId);
			}
			else {
				return JSON.stringify(value);
			}
		};
		RequestPrettyPrinter.prototype.getObjVarNameBase=function (objectPathId) {
			var ret="v";
			var objPath=this.m_referencedObjectPaths[objectPathId];
			if (objPath) {
				switch (objPath.objectPathInfo.ObjectPathType) {
					case 1:
						ret=this.m_globalObjName;
						break;
					case 4:
						ret=OfficeExtension.Utility._toCamelLowerCase(objPath.objectPathInfo.Name);
						break;
					case 3:
						var methodName=objPath.objectPathInfo.Name;
						if (methodName.length > 3 && methodName.substr(0, 3)==="Get") {
							methodName=methodName.substr(3);
						}
						ret=OfficeExtension.Utility._toCamelLowerCase(methodName);
						break;
					case 5:
						var parentName=this.getObjVarNameBase(objPath.objectPathInfo.ParentObjectPathId);
						if (parentName.charAt(parentName.length - 1)==="s") {
							ret=parentName.substr(0, parentName.length - 1);
						}
						else {
							ret=parentName+"Item";
						}
						break;
				}
			}
			return ret;
		};
		RequestPrettyPrinter.prototype.getObjVarName=function (objectPathId) {
			if (this.m_variableNameForObjectPathMap[objectPathId]) {
				return this.m_variableNameForObjectPathMap[objectPathId];
			}
			var ret=this.getObjVarNameBase(objectPathId);
			if (!this.m_variableNameToObjectPathMap[ret]) {
				this.m_variableNameForObjectPathMap[objectPathId]=ret;
				this.m_variableNameToObjectPathMap[ret]=objectPathId;
				return ret;
			}
			var i=1;
			while (this.m_variableNameToObjectPathMap[ret+i.toString()]) {
				i++;
			}
			ret=ret+i.toString();
			this.m_variableNameForObjectPathMap[objectPathId]=ret;
			this.m_variableNameToObjectPathMap[ret]=objectPathId;
			return ret;
		};
		return RequestPrettyPrinter;
	}());
	OfficeExtension.RequestPrettyPrinter=RequestPrettyPrinter;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ResourceStrings=(function () {
		function ResourceStrings() {
		}
		ResourceStrings.cannotRegisterEvent="CannotRegisterEvent";
		ResourceStrings.connectionFailureWithStatus="ConnectionFailureWithStatus";
		ResourceStrings.connectionFailureWithDetails="ConnectionFailureWithDetails";
		ResourceStrings.invalidObjectPath="InvalidObjectPath";
		ResourceStrings.invalidRequestContext="InvalidRequestContext";
		ResourceStrings.invalidArgument="InvalidArgument";
		ResourceStrings.invalidArgumentGeneric="InvalidArgumentGeneric";
		ResourceStrings.propertyNotLoaded="PropertyNotLoaded";
		ResourceStrings.runMustReturnPromise="RunMustReturnPromise";
		ResourceStrings.timeout="Timeout";
		ResourceStrings.propertyDoesNotExist="PropertyDoesNotExist";
		ResourceStrings.attemptingToSetReadOnlyProperty="AttemptingToSetReadOnlyProperty";
		ResourceStrings.moreInfoInnerError="MoreInfoInnerError";
		ResourceStrings.cannotApplyPropertyThroughSetMethod="CannotApplyPropertyThroughSetMethod";
		ResourceStrings.valueNotLoaded="ValueNotLoaded";
		ResourceStrings.invalidOrTimedOutSessionMessage="InvalidOrTimedOutSessionMessage";
		ResourceStrings.invalidOperationInCellEditMode="InvalidOperationInCellEditMode";
		ResourceStrings.customFunctionDefintionMissing="CustomFunctionDefintionMissing";
		ResourceStrings.customFunctionImplementationMissing="CustomFunctionImplementationMissing";
		ResourceStrings.customFunctionNameContainsBadChars="CustomFunctionNameContainsBadChars";
		ResourceStrings.customFunctionNameCannotSplit="CustomFunctionNameCannotSplit";
		ResourceStrings.customFunctionUnexpectedNumberOfEntriesInResultBatch="CustomFunctionUnexpectedNumberOfEntriesInResultBatch";
		ResourceStrings.customFunctionCancellationHandlerMissing="CustomFunctionCancellationHandlerMissing";
		ResourceStrings.apiNotFoundDetails="ApiNotFoundDetails";
		ResourceStrings.pendingBatchInProgress="PendingBatchInProgress";
		ResourceStrings.notInsideBatch="NotInsideBatch";
		ResourceStrings.cannotUpdateReadOnlyProperty="CannotUpdateReadOnlyProperty";
		return ResourceStrings;
	}());
	OfficeExtension.ResourceStrings=ResourceStrings;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var ResourceStringValues=(function () {
		function ResourceStringValues() {
		}
		ResourceStringValues.CannotRegisterEvent="The event handler cannot be registered.";
		ResourceStringValues.ConnectionFailureWithStatus="The request failed with status code of {0}.";
		ResourceStringValues.ConnectionFailureWithDetails="The request failed with status code of {0}, error code {1} and the following error message: {2}";
		ResourceStringValues.InvalidArgument="The argument '{0}' doesn't work for this situation, is missing, or isn't in the right format.";
		ResourceStringValues.InvalidObjectPath="The object path '{0}' isn't working for what you're trying to do. If you're using the object across multiple \"context.sync\" calls and outside the sequential execution of a \".run\" batch, please use the \"context.trackedObjects.add()\" and \"context.trackedObjects.remove()\" methods to manage the object's lifetime.";
		ResourceStringValues.InvalidRequestContext="Cannot use the object across different request contexts.";
		ResourceStringValues.PropertyNotLoaded="The property '{0}' is not available. Before reading the property's value, call the load method on the containing object and call \"context.sync()\" on the associated request context.";
		ResourceStringValues.RunMustReturnPromise="The batch function passed to the \".run\" method didn't return a promise. The function must return a promise, so that any automatically-tracked objects can be released at the completion of the batch operation. Typically, you return a promise by returning the response from \"context.sync()\".";
		ResourceStringValues.Timeout="The operation has timed out.";
		ResourceStringValues.ValueNotLoaded="The value of the result object has not been loaded yet. Before reading the value property, call \"context.sync()\" on the associated request context.";
		ResourceStringValues.InvalidOrTimedOutSessionMessage="Your Office Online session has expired or is invalid. To continue, refresh the page.";
		ResourceStringValues.InvalidOperationInCellEditMode="Excel is in cell-editing mode. Please exit the edit mode by pressing ENTER or TAB or selecting another cell, and then try again.";
		ResourceStringValues.CustomFunctionDefintionMissing="A property with this name that represents the function's definition must exist on Excel.CustomFunctions.";
		ResourceStringValues.CustomFunctionImplementationMissing="The property with this name on Excel.CustomFunctions that represents the function's definition must contain a 'call' property that implements the function.";
		ResourceStringValues.CustomFunctionNameContainsBadChars="The function name may only contain letters, digits, underscores, and periods.";
		ResourceStringValues.CustomFunctionNameCannotSplit="The function name must contain a non-empty namespace and a non-empty short name.";
		ResourceStringValues.CustomFunctionUnexpectedNumberOfEntriesInResultBatch="The batching function returned a number of results that doesn't match the number of parameter value sets that were passed into it.";
		ResourceStringValues.CustomFunctionCancellationHandlerMissing="The cancellation handler onCanceled is missing in the function. The handler must be present as the function is defined as cancelable.";
		ResourceStringValues.ApiNotFoundDetails="The method or property {0} is part of the {1} requirement set, which is not available in your version of {2}.";
		ResourceStringValues.PendingBatchInProgress="There is a pending batch in progress. The batch method may not be called inside another batch, or simultaneously with another batch.";
		ResourceStringValues.NotInsideBatch="Operations may not be invoked outside of a batch method.";
		ResourceStringValues.CannotUpdateReadOnlyProperty="The property '{0}' is read-only and it cannot be updated.";
		return ResourceStringValues;
	}());
	OfficeExtension.ResourceStringValues=ResourceStringValues;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var RichApiMessageUtility=(function () {
		function RichApiMessageUtility() {
		}
		RichApiMessageUtility.buildMessageArrayForIRequestExecutor=function (customData, requestFlags, requestMessage, sourceLibHeaderValue) {
			var requestMessageText=JSON.stringify(requestMessage.Body);
			OfficeExtension.Utility.log("Request:");
			OfficeExtension.Utility.log(requestMessageText);
			var headers={};
			headers[OfficeExtension.Constants.sourceLibHeader]=sourceLibHeaderValue;
			var messageSafearray=RichApiMessageUtility.buildRequestMessageSafeArray(customData, requestFlags, "POST", "ProcessQuery", headers, requestMessageText);
			return messageSafearray;
		};
		RichApiMessageUtility.buildResponseOnSuccess=function (responseBody, responseHeaders) {
			var response={ ErrorCode: '', ErrorMessage: '', Headers: null, Body: null };
			response.Body=JSON.parse(responseBody);
			response.Headers=responseHeaders;
			return response;
		};
		RichApiMessageUtility.buildResponseOnError=function (errorCode, message) {
			var response={ ErrorCode: '', ErrorMessage: '', Headers: null, Body: null };
			response.ErrorCode=OfficeExtension.ErrorCodes.generalException;
			response.ErrorMessage=message;
			if (errorCode==RichApiMessageUtility.OfficeJsErrorCode_ooeNoCapability) {
				response.ErrorCode=OfficeExtension.ErrorCodes.accessDenied;
			}
			else if (errorCode==RichApiMessageUtility.OfficeJsErrorCode_ooeActivityLimitReached) {
				response.ErrorCode=OfficeExtension.ErrorCodes.activityLimitReached;
			}
			else if (errorCode==RichApiMessageUtility.OfficeJsErrorCode_ooeInvalidOrTimedOutSession) {
				response.ErrorCode=OfficeExtension.ErrorCodes.invalidOrTimedOutSession;
				response.ErrorMessage=OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.invalidOrTimedOutSessionMessage);
			}
			return response;
		};
		RichApiMessageUtility.buildHttpResponseFromOfficeJsError=function (errorCode, message) {
			var statusCode=500;
			var errorBody={};
			errorBody["error"]={};
			errorBody["error"]["code"]=OfficeExtension.ErrorCodes.generalException;
			errorBody["error"]["message"]=message;
			if (errorCode===RichApiMessageUtility.OfficeJsErrorCode_ooeNoCapability) {
				statusCode=403;
				errorBody["error"]["code"]=OfficeExtension.ErrorCodes.accessDenied;
			}
			else if (errorCode===RichApiMessageUtility.OfficeJsErrorCode_ooeActivityLimitReached) {
				statusCode=429;
				errorBody["error"]["code"]=OfficeExtension.ErrorCodes.activityLimitReached;
			}
			return { statusCode: statusCode, headers: {}, body: JSON.stringify(errorBody) };
		};
		RichApiMessageUtility.buildRequestMessageSafeArray=function (customData, requestFlags, method, path, headers, body) {
			var headerArray=[];
			if (headers) {
				for (var headerName in headers) {
					headerArray.push(headerName);
					headerArray.push(headers[headerName]);
				}
			}
			var appPermission=0;
			var solutionId="";
			var instanceId="";
			var marketplaceType="";
			return [
				customData,
				method,
				path,
				headerArray,
				body,
				appPermission,
				requestFlags,
				solutionId,
				instanceId,
				marketplaceType
			];
		};
		RichApiMessageUtility.getResponseBody=function (result) {
			return RichApiMessageUtility.getResponseBodyFromSafeArray(result.value.data);
		};
		RichApiMessageUtility.getResponseHeaders=function (result) {
			return RichApiMessageUtility.getResponseHeadersFromSafeArray(result.value.data);
		};
		RichApiMessageUtility.getResponseBodyFromSafeArray=function (data) {
			var ret=data[2];
			if (typeof (ret)==="string") {
				return ret;
			}
			var arr=ret;
			return arr.join("");
		};
		RichApiMessageUtility.getResponseHeadersFromSafeArray=function (data) {
			var arrayHeader=data[1];
			if (!arrayHeader) {
				return null;
			}
			var headers={};
			for (var i=0; i < arrayHeader.length - 1; i+=2) {
				headers[arrayHeader[i]]=arrayHeader[i+1];
			}
			return headers;
		};
		RichApiMessageUtility.getResponseStatusCode=function (result) {
			return RichApiMessageUtility.getResponseStatusCodeFromSafeArray(result.value.data);
		};
		RichApiMessageUtility.getResponseStatusCodeFromSafeArray=function (data) {
			return data[0];
		};
		RichApiMessageUtility.OfficeJsErrorCode_ooeInvalidOrTimedOutSession=5012;
		RichApiMessageUtility.OfficeJsErrorCode_ooeActivityLimitReached=5102;
		RichApiMessageUtility.OfficeJsErrorCode_ooeNoCapability=7000;
		return RichApiMessageUtility;
	}());
	OfficeExtension.RichApiMessageUtility=RichApiMessageUtility;
})(OfficeExtension || (OfficeExtension={}));
var OfficeExtension;
(function (OfficeExtension) {
	var Utility=(function () {
		function Utility() {
		}
		Utility.checkArgumentNull=function (value, name) {
			if (Utility.isNullOrUndefined(value)) {
				throw OfficeExtension._Internal.RuntimeError._createInvalidArgError(name);
			}
		};
		Utility.isNullOrUndefined=function (value) {
			if (value===null) {
				return true;
			}
			if (typeof (value)==="undefined") {
				return true;
			}
			return false;
		};
		Utility.isUndefined=function (value) {
			if (typeof (value)==="undefined") {
				return true;
			}
			return false;
		};
		Utility.isNullOrEmptyString=function (value) {
			if (value===null) {
				return true;
			}
			if (typeof (value)==="undefined") {
				return true;
			}
			if (value.length==0) {
				return true;
			}
			return false;
		};
		Utility.isPlainJsonObject=function (value) {
			if (Utility.isNullOrUndefined(value)) {
				return false;
			}
			if (typeof (value) !=="object") {
				return false;
			}
			return Object.getPrototypeOf(value)===Object.getPrototypeOf({});
		};
		Utility.trim=function (str) {
			return str.replace(new RegExp("^\\s+|\\s+$", "g"), "");
		};
		Utility.caseInsensitiveCompareString=function (str1, str2) {
			if (Utility.isNullOrUndefined(str1)) {
				return Utility.isNullOrUndefined(str2);
			}
			else {
				if (Utility.isNullOrUndefined(str2)) {
					return false;
				}
				else {
					return str1.toUpperCase()==str2.toUpperCase();
				}
			}
		};
		Utility.adjustToDateTime=function (value) {
			if (Utility.isNullOrUndefined(value)) {
				return null;
			}
			if (typeof (value)==="string") {
				return new Date(value);
			}
			if (Array.isArray(value)) {
				var arr=value;
				for (var i=0; i < arr.length; i++) {
					arr[i]=Utility.adjustToDateTime(arr[i]);
				}
				return arr;
			}
			throw OfficeExtension._Internal.RuntimeError._createInvalidArgError("date");
		};
		Utility.isReadonlyRestRequest=function (method) {
			return Utility.caseInsensitiveCompareString(method, "GET");
		};
		Utility.setMethodArguments=function (context, argumentInfo, args) {
			if (Utility.isNullOrUndefined(args)) {
				return null;
			}
			var referencedObjectPaths=new Array();
			var referencedObjectPathIds=new Array();
			var hasOne=Utility.collectObjectPathInfos(context, args, referencedObjectPaths, referencedObjectPathIds);
			argumentInfo.Arguments=args;
			if (hasOne) {
				argumentInfo.ReferencedObjectPathIds=referencedObjectPathIds;
				return referencedObjectPaths;
			}
			return null;
		};
		Utility.collectObjectPathInfos=function (context, args, referencedObjectPaths, referencedObjectPathIds) {
			var hasOne=false;
			for (var i=0; i < args.length; i++) {
				if (args[i] instanceof OfficeExtension.ClientObject) {
					var clientObject=args[i];
					Utility.validateContext(context, clientObject);
					args[i]=clientObject._objectPath.objectPathInfo.Id;
					referencedObjectPathIds.push(clientObject._objectPath.objectPathInfo.Id);
					referencedObjectPaths.push(clientObject._objectPath);
					hasOne=true;
				}
				else if (Array.isArray(args[i])) {
					var childArrayObjectPathIds=new Array();
					var childArrayHasOne=Utility.collectObjectPathInfos(context, args[i], referencedObjectPaths, childArrayObjectPathIds);
					if (childArrayHasOne) {
						referencedObjectPathIds.push(childArrayObjectPathIds);
						hasOne=true;
					}
					else {
						referencedObjectPathIds.push(0);
					}
				}
				else {
					referencedObjectPathIds.push(0);
				}
			}
			return hasOne;
		};
		Utility.fixObjectPathIfNecessary=function (clientObject, value) {
			if (clientObject && clientObject._objectPath && value) {
				clientObject._objectPath.updateUsingObjectData(value, clientObject);
			}
		};
		Utility.tryGetObjectIdFromLoadOrRetrieveResult=function (value) {
			var id=value[OfficeExtension.Constants.id];
			if (Utility.isNullOrUndefined(id)) {
				id=value[OfficeExtension.Constants.idLowerCase];
			}
			if (Utility.isNullOrUndefined(id)) {
				id=value[OfficeExtension.Constants.idPrivate];
			}
			return id;
		};
		Utility.validateObjectPath=function (clientObject) {
			var objectPath=clientObject._objectPath;
			while (objectPath) {
				if (!objectPath.isValid) {
					throw new OfficeExtension._Internal.RuntimeError({
						code: OfficeExtension.ErrorCodes.invalidObjectPath,
						message: Utility._getResourceString(OfficeExtension.ResourceStrings.invalidObjectPath, Utility.getObjectPathExpression(objectPath)),
						debugInfo: {
							errorLocation: Utility.getObjectPathExpression(objectPath)
						}
					});
				}
				objectPath=objectPath.parentObjectPath;
			}
		};
		Utility.validateReferencedObjectPaths=function (objectPaths) {
			if (objectPaths) {
				for (var i=0; i < objectPaths.length; i++) {
					var objectPath=objectPaths[i];
					while (objectPath) {
						if (!objectPath.isValid) {
							throw new OfficeExtension._Internal.RuntimeError({
								code: OfficeExtension.ErrorCodes.invalidObjectPath,
								message: Utility._getResourceString(OfficeExtension.ResourceStrings.invalidObjectPath, Utility.getObjectPathExpression(objectPath))
							});
						}
						objectPath=objectPath.parentObjectPath;
					}
				}
			}
		};
		Utility.validateContext=function (context, obj) {
			if (obj && obj.context !==context) {
				throw new OfficeExtension._Internal.RuntimeError({
					code: OfficeExtension.ErrorCodes.invalidRequestContext,
					message: Utility._getResourceString(OfficeExtension.ResourceStrings.invalidRequestContext)
				});
			}
		};
		Utility.log=function (message) {
			if (Utility._logEnabled && typeof (console) !=="undefined" && console.log) {
				console.log(message);
			}
		};
		Utility.load=function (clientObj, option) {
			clientObj.context.load(clientObj, option);
			return clientObj;
		};
		Utility.loadAndSync=function (clientObj, option) {
			clientObj.context.load(clientObj, option);
			return clientObj.context.sync().then(function () { return clientObj; });
		};
		Utility.retrieve=function (clientObj, option) {
			var shouldPolyfill=OfficeExtension._internalConfig.alwaysPolyfillClientObjectRetrieveMethod;
			if (!shouldPolyfill) {
				shouldPolyfill=!Utility.isSetSupported("RichApiRuntime", "1.1");
			}
			var result=new OfficeExtension.RetrieveResultImpl(clientObj, shouldPolyfill);
			var queryOption=OfficeExtension.ClientRequestContext._parseQueryOption(option);
			var action;
			if (shouldPolyfill) {
				action=OfficeExtension.ActionFactory.createQueryAction(clientObj.context, clientObj, queryOption);
			}
			else {
				action=OfficeExtension.ActionFactory.createQueryAsJsonAction(clientObj.context, clientObj, queryOption);
			}
			clientObj.context._pendingRequest.addActionResultHandler(action, result);
			return result;
		};
		Utility.retrieveAndSync=function (clientObj, option) {
			var result=Utility.retrieve(clientObj, option);
			return clientObj.context.sync().then(function () { return result; });
		};
		Utility.isSetSupported=function (apiSetName, apiSetVersion) {
			if (typeof (window) !=="undefined" && window.Office && window.Office.context && window.Office.context.requirements) {
				return window.Office.context.requirements.isSetSupported(apiSetName, apiSetVersion);
			}
			return true;
		};
		Utility._parseSelectExpand=function (select) {
			var args=[];
			if (!Utility.isNullOrEmptyString(select)) {
				var propertyNames=select.split(",");
				for (var i=0; i < propertyNames.length; i++) {
					var propertyName=propertyNames[i];
					propertyName=sanitizeForAnyItemsSlash(propertyName.trim());
					if (propertyName.length > 0) {
						args.push(propertyName);
					}
				}
			}
			return args;
			function sanitizeForAnyItemsSlash(propertyName) {
				var propertyNameLower=propertyName.toLowerCase();
				if (propertyNameLower==="items" || propertyNameLower==="items/") {
					return '*';
				}
				var itemsSlashLength=6;
				var isItemsSlashOrItemsDot=propertyNameLower.substr(0, itemsSlashLength)==="items/" ||
					propertyNameLower.substr(0, itemsSlashLength)==="items.";
				if (isItemsSlashOrItemsDot) {
					propertyName=propertyName.substr(itemsSlashLength);
				}
				return propertyName.replace(new RegExp("[\/\.]items[\/\.]", "gi"), "/");
			}
		};
		Utility.toJson=function (clientObj, scalarProperties, navigationProperties, collectionItemsIfAny) {
			var result={};
			for (var prop in scalarProperties) {
				var value=scalarProperties[prop];
				if (typeof value !=="undefined") {
					result[prop]=value;
				}
			}
			for (var prop in navigationProperties) {
				var value=navigationProperties[prop];
				if (typeof value !=="undefined") {
					if (value[Utility.fieldName_isCollection] && (typeof value[Utility.fieldName_m__items] !=="undefined")) {
						result[prop]=value.toJSON()["items"];
					}
					else {
						result[prop]=value.toJSON();
					}
				}
			}
			if (collectionItemsIfAny) {
				result["items"]=collectionItemsIfAny.map(function (item) { return item.toJSON(); });
			}
			return result;
		};
		Utility.throwError=function (resourceId, arg, errorLocation) {
			throw new OfficeExtension._Internal.RuntimeError({
				code: resourceId,
				message: Utility._getResourceString(resourceId, arg),
				debugInfo: errorLocation ? { errorLocation: errorLocation } : undefined
			});
		};
		Utility.createRuntimeError=function (code, message, location) {
			return (new OfficeExtension._Internal.RuntimeError({
				code: code,
				message: message,
				debugInfo: { errorLocation: location }
			}));
		};
		Utility._getResourceString=function (resourceId, arg) {
			var ret;
			if (typeof (window) !=="undefined" && window.Strings && window.Strings.OfficeOM) {
				var stringName="L_"+resourceId;
				var stringValue=window.Strings.OfficeOM[stringName];
				if (stringValue) {
					ret=stringValue;
				}
			}
			if (!ret) {
				ret=OfficeExtension.ResourceStringValues[resourceId];
			}
			if (!ret) {
				ret=resourceId;
			}
			if (!Utility.isNullOrUndefined(arg)) {
				if (Array.isArray(arg)) {
					var arrArg=arg;
					ret=Utility._formatString(ret, arrArg);
				}
				else {
					ret=ret.replace("{0}", arg);
				}
			}
			return ret;
		};
		Utility._formatString=function (format, arrArg) {
			return format.replace(/\{\d\}/g, function (v) {
				var position=parseInt(v.substr(1, v.length - 2));
				if (position < arrArg.length) {
					return arrArg[position];
				}
				else {
					throw OfficeExtension._Internal.RuntimeError._createInvalidArgError("format");
				}
			});
		};
		Utility.throwIfNotLoaded=function (propertyName, fieldValue, entityName, isNull) {
			if (!isNull && Utility.isUndefined(fieldValue) && propertyName.charCodeAt(0) !=Utility.s_underscoreCharCode) {
				throw Utility.createPropertyNotLoadedException(entityName, propertyName);
			}
		};
		Utility.createPropertyNotLoadedException=function (entityName, propertyName) {
			return new OfficeExtension._Internal.RuntimeError({
				code: OfficeExtension.ErrorCodes.propertyNotLoaded,
				message: Utility._getResourceString(OfficeExtension.ResourceStrings.propertyNotLoaded, propertyName),
				debugInfo: entityName ? { errorLocation: entityName+"."+propertyName } : undefined
			});
		};
		Utility.createCannotUpdateReadOnlyPropertyException=function (entityName, propertyName) {
			return new OfficeExtension._Internal.RuntimeError({
				code: OfficeExtension.ErrorCodes.cannotUpdateReadOnlyProperty,
				message: Utility._getResourceString(OfficeExtension.ResourceStrings.cannotUpdateReadOnlyProperty, propertyName),
				debugInfo: entityName ? { errorLocation: entityName+"."+propertyName } : undefined
			});
		};
		Utility.throwIfApiNotSupported=function (apiFullName, apiSetName, apiSetVersion, hostName) {
			if (!Utility._doApiNotSupportedCheck) {
				return;
			}
			if (!Utility.isSetSupported(apiSetName, apiSetVersion)) {
				var message=Utility._getResourceString(OfficeExtension.ResourceStrings.apiNotFoundDetails, [apiFullName, apiSetName+" "+apiSetVersion, hostName]);
				throw new OfficeExtension._Internal.RuntimeError({
					code: OfficeExtension.ErrorCodes.apiNotFound,
					message: message,
					debugInfo: { errorLocation: apiFullName }
				});
			}
		};
		Utility.getObjectPathExpression=function (objectPath) {
			var ret="";
			while (objectPath) {
				switch (objectPath.objectPathInfo.ObjectPathType) {
					case 1:
						ret=ret;
						break;
					case 2:
						ret="new()"+(ret.length > 0 ? "." : "")+ret;
						break;
					case 3:
						ret=Utility.normalizeName(objectPath.objectPathInfo.Name)+"()"+(ret.length > 0 ? "." : "")+ret;
						break;
					case 4:
						ret=Utility.normalizeName(objectPath.objectPathInfo.Name)+(ret.length > 0 ? "." : "")+ret;
						break;
					case 5:
						ret="getItem()"+(ret.length > 0 ? "." : "")+ret;
						break;
					case 6:
						ret="_reference()"+(ret.length > 0 ? "." : "")+ret;
						break;
				}
				objectPath=objectPath.parentObjectPath;
			}
			return ret;
		};
		Utility._createPromiseFromResult=function (value) {
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				resolve(value);
			});
		};
		Utility._createTimeoutPromise=function (timeout) {
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				setTimeout(function () {
					resolve(null);
				}, timeout);
			});
		};
		Utility.promisify=function (action) {
			return new OfficeExtension._Internal.OfficePromise(function (resolve, reject) {
				var callback=function (result) {
					if (result.status=="failed") {
						reject(result.error);
					}
					else {
						resolve(result.value);
					}
				};
				action(callback);
			});
		};
		Utility._addActionResultHandler=function (clientObj, action, resultHandler) {
			clientObj.context._pendingRequest.addActionResultHandler(action, resultHandler);
		};
		Utility._handleNavigationPropertyResults=function (clientObj, objectValue, propertyNames) {
			for (var i=0; i < propertyNames.length - 1; i+=2) {
				if (!Utility.isUndefined(objectValue[propertyNames[i+1]])) {
					clientObj[propertyNames[i]]._handleResult(objectValue[propertyNames[i+1]]);
				}
			}
		};
		Utility.normalizeName=function (name) {
			return name.substr(0, 1).toLowerCase()+name.substr(1);
		};
		Utility._isLocalDocumentUrl=function (url) {
			return Utility._getLocalDocumentUrlPrefixLength(url) > 0;
		};
		Utility._getLocalDocumentUrlPrefixLength=function (url) {
			var localDocumentPrefixes=["http://document.localhost", "https://document.localhost", "//document.localhost"];
			var urlLower=url.toLowerCase().trim();
			for (var i=0; i < localDocumentPrefixes.length; i++) {
				if (urlLower===localDocumentPrefixes[i]) {
					return localDocumentPrefixes[i].length;
				}
				else if (urlLower.substr(0, localDocumentPrefixes[i].length+1)===localDocumentPrefixes[i]+"/") {
					return localDocumentPrefixes[i].length+1;
				}
			}
			return 0;
		};
		Utility._validateLocalDocumentRequest=function (request) {
			var index=Utility._getLocalDocumentUrlPrefixLength(request.url);
			if (index <=0) {
				throw OfficeExtension._Internal.RuntimeError._createInvalidArgError({
					argumentName: "request"
				});
			}
			var path=request.url.substr(index);
			var pathLower=path.toLowerCase();
			if (pathLower==="_api") {
				path="";
			}
			else if (pathLower.substr(0, "_api/".length)==="_api/") {
				path=path.substr("_api/".length);
			}
			return {
				method: request.method,
				url: path,
				headers: request.headers,
				body: request.body
			};
		};
		Utility._buildRequestMessageSafeArray=function (request) {
			var requestFlags=0;
			if (!Utility.isReadonlyRestRequest(request.method)) {
				requestFlags=1;
			}
			if (request.url.substr(0, OfficeExtension.Constants.processQuery.length).toLowerCase()===OfficeExtension.Constants.processQuery.toLowerCase()) {
				var index=request.url.indexOf("?");
				if (index > 0) {
					var queryString=request.url.substr(index+1);
					var parts=queryString.split("&");
					for (var i=0; i < parts.length; i++) {
						var keyvalue=parts[i].split("=");
						if (keyvalue[0].toLowerCase()===OfficeExtension.Constants.flags) {
							var flags=parseInt(keyvalue[1]);
							requestFlags=flags;
							requestFlags=requestFlags & 1;
							break;
						}
					}
				}
			}
			return OfficeExtension.RichApiMessageUtility.buildRequestMessageSafeArray("", requestFlags, request.method, request.url, request.headers, request.body);
		};
		Utility._parseHttpResponseHeaders=function (allResponseHeaders) {
			var responseHeaders={};
			if (!Utility.isNullOrEmptyString(allResponseHeaders)) {
				var regex=new RegExp("\r?\n");
				var entries=allResponseHeaders.split(regex);
				for (var i=0; i < entries.length; i++) {
					var entry=entries[i];
					if (entry !=null) {
						var index=entry.indexOf(':');
						if (index > 0) {
							var key=entry.substr(0, index);
							var value=entry.substr(index+1);
							key=Utility.trim(key);
							value=Utility.trim(value);
							responseHeaders[key.toUpperCase()]=value;
						}
					}
				}
			}
			return responseHeaders;
		};
		Utility._parseErrorResponse=function (responseInfo) {
			var errorObj=null;
			if (Utility.isPlainJsonObject(responseInfo.body)) {
				errorObj=responseInfo.body;
			}
			else if (!Utility.isNullOrEmptyString(responseInfo.body)) {
				var errorResponseBody=Utility.trim(responseInfo.body);
				try {
					errorObj=JSON.parse(errorResponseBody);
				}
				catch (e) {
					Utility.log("Error when parse "+errorResponseBody);
				}
			}
			var errorMessage;
			var errorCode;
			if (!Utility.isNullOrUndefined(errorObj) && typeof (errorObj)==="object" && errorObj.error) {
				errorCode=errorObj.error.code;
				errorMessage=Utility._getResourceString(OfficeExtension.ResourceStrings.connectionFailureWithDetails, [responseInfo.statusCode.toString(), errorObj.error.code, errorObj.error.message]);
			}
			else {
				errorMessage=Utility._getResourceString(OfficeExtension.ResourceStrings.connectionFailureWithStatus, responseInfo.statusCode.toString());
			}
			if (Utility.isNullOrEmptyString(errorCode)) {
				errorCode=OfficeExtension.ErrorCodes.connectionFailure;
			}
			return { errorCode: errorCode, errorMessage: errorMessage };
		};
		Utility._copyHeaders=function (src, dest) {
			if (src && dest) {
				for (var key in src) {
					dest[key]=src[key];
				}
			}
		};
		Utility._toCamelLowerCase=function (name) {
			if (Utility.isNullOrEmptyString(name)) {
				return name;
			}
			var index=0;
			while (index < name.length && name.charCodeAt(index) >=65 && name.charCodeAt(index) <=90) {
				index++;
			}
			if (index < name.length) {
				return name.substr(0, index).toLowerCase()+name.substr(index);
			}
			else {
				return name.toLowerCase();
			}
		};
		Utility.definePropertyThrowUnloadedException=function (obj, typeName, propertyName) {
			Object.defineProperty(obj, propertyName, {
				configurable: true,
				enumerable: true,
				get: function () {
					throw Utility.createPropertyNotLoadedException(typeName, propertyName);
				},
				set: function () {
					throw Utility.createCannotUpdateReadOnlyPropertyException(typeName, propertyName);
				}
			});
		};
		Utility.defineReadOnlyPropertyWithValue=function (obj, propertyName, value) {
			Object.defineProperty(obj, propertyName, {
				configurable: true,
				enumerable: true,
				get: function () {
					return value;
				},
				set: function () {
					throw Utility.createCannotUpdateReadOnlyPropertyException(null, propertyName);
				}
			});
		};
		Utility.processRetrieveResult=function (proxy, value, result, childItemCreateFunc) {
			if (Utility.isNullOrUndefined(value)) {
				return;
			}
			if (childItemCreateFunc) {
				var data=value[OfficeExtension.Constants.itemsLowerCase];
				if (Array.isArray(data)) {
					var itemsResult=[];
					for (var i=0; i < data.length; i++) {
						var itemProxy=childItemCreateFunc(data[i], i);
						var itemResult={};
						itemResult[OfficeExtension.Constants.proxy]=itemProxy;
						itemProxy._handleRetrieveResult(data[i], itemResult);
						itemsResult.push(itemResult);
					}
					Utility.defineReadOnlyPropertyWithValue(result, OfficeExtension.Constants.itemsLowerCase, itemsResult);
				}
			}
			else {
				var scalarPropertyNames=proxy[OfficeExtension.Constants.scalarPropertyNames];
				var navigationPropertyNames=proxy[OfficeExtension.Constants.navigationPropertyNames];
				var typeName=proxy[OfficeExtension.Constants.className];
				if (scalarPropertyNames) {
					for (var i=0; i < scalarPropertyNames.length; i++) {
						var propName=scalarPropertyNames[i];
						var propValue=value[propName];
						if (Utility.isUndefined(propValue)) {
							Utility.definePropertyThrowUnloadedException(result, typeName, propName);
						}
						else {
							Utility.defineReadOnlyPropertyWithValue(result, propName, propValue);
						}
					}
				}
				if (navigationPropertyNames) {
					for (var i=0; i < navigationPropertyNames.length; i++) {
						var propName=navigationPropertyNames[i];
						var propValue=value[propName];
						if (Utility.isUndefined(propValue)) {
							Utility.definePropertyThrowUnloadedException(result, typeName, propName);
						}
						else {
							var propProxy=proxy[propName];
							var propResult={};
							propProxy._handleRetrieveResult(propValue, propResult);
							propResult[OfficeExtension.Constants.proxy]=propProxy;
							if (Array.isArray(propResult[OfficeExtension.Constants.itemsLowerCase])) {
								propResult=propResult[OfficeExtension.Constants.itemsLowerCase];
							}
							Utility.defineReadOnlyPropertyWithValue(result, propName, propResult);
						}
					}
				}
			}
		};
		Utility.fieldName_m__items="m__items";
		Utility.fieldName_isCollection="_isCollection";
		Utility._logEnabled=false;
		Utility._synchronousCleanup=false;
		Utility._doApiNotSupportedCheck=false;
		Utility.s_underscoreCharCode="_".charCodeAt(0);
		return Utility;
	}());
	OfficeExtension.Utility=Utility;
})(OfficeExtension || (OfficeExtension={}));

var __extends=(this && this.__extends) || (function () {
	var extendStatics=Object.setPrototypeOf ||
		({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__=b; }) ||
		function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p]=b[p]; };
	return function (d, b) {
		extendStatics(d, b);
		function __() { this.constructor=d; }
		d.prototype=b===null ? Object.create(b) : (__.prototype=b.prototype, new __());
	};
})();
var OfficeCore;
(function (OfficeCore) {
	var _hostName="OfficeCore";
	var _defaultApiSetName="AgaveVisualApi";
	var _createPropertyObjectPath=OfficeExtension.ObjectPathFactory.createPropertyObjectPath;
	var _createMethodObjectPath=OfficeExtension.ObjectPathFactory.createMethodObjectPath;
	var _createIndexerObjectPath=OfficeExtension.ObjectPathFactory.createIndexerObjectPath;
	var _createNewObjectObjectPath=OfficeExtension.ObjectPathFactory.createNewObjectObjectPath;
	var _createChildItemObjectPathUsingIndexer=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexer;
	var _createChildItemObjectPathUsingGetItemAt=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingGetItemAt;
	var _createChildItemObjectPathUsingIndexerOrGetItemAt=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt;
	var _createMethodAction=OfficeExtension.ActionFactory.createMethodAction;
	var _createEnsureUnchangedAction=OfficeExtension.ActionFactory.createEnsureUnchangedAction;
	var _createSetPropertyAction=OfficeExtension.ActionFactory.createSetPropertyAction;
	var _isNullOrUndefined=OfficeExtension.Utility.isNullOrUndefined;
	var _isUndefined=OfficeExtension.Utility.isUndefined;
	var _throwIfNotLoaded=OfficeExtension.Utility.throwIfNotLoaded;
	var _throwIfApiNotSupported=OfficeExtension.Utility.throwIfApiNotSupported;
	var _load=OfficeExtension.Utility.load;
	var _loadAndSync=OfficeExtension.Utility.loadAndSync;
	var _retrieve=OfficeExtension.Utility.retrieve;
	var _retrieveAndSync=OfficeExtension.Utility.retrieveAndSync;
	var _toJson=OfficeExtension.Utility.toJson;
	var _fixObjectPathIfNecessary=OfficeExtension.Utility.fixObjectPathIfNecessary;
	var _addActionResultHandler=OfficeExtension.Utility._addActionResultHandler;
	var _handleNavigationPropertyResults=OfficeExtension.Utility._handleNavigationPropertyResults;
	var _adjustToDateTime=OfficeExtension.Utility.adjustToDateTime;
	var _typeBiShim="BiShim";
	var BiShim=(function (_super) {
		__extends(BiShim, _super);
		function BiShim() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(BiShim.prototype, "_className", {
			get: function () {
				return "BiShim";
			},
			enumerable: true,
			configurable: true
		});
		BiShim.prototype.initialize=function (capabilities) {
			_createMethodAction(this.context, this, "Initialize", 0, [capabilities], false);
		};
		BiShim.prototype.uninitialize=function () {
			_createMethodAction(this.context, this, "Uninitialize", 0, [], false);
		};
		BiShim.prototype.getData=function () {
			var action=_createMethodAction(this.context, this, "getData", 1, [], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		BiShim.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
		};
		BiShim.newObject=function (context) {
			var ret=new OfficeCore.BiShim(context, _createNewObjectObjectPath(context, "Microsoft.AgaveVisual.BiShim", false, false));
			return ret;
		};
		BiShim.prototype.toJSON=function () {
			return _toJson(this, {}, {});
		};
		return BiShim;
	}(OfficeExtension.ClientObject));
	OfficeCore.BiShim=BiShim;
	var ErrorCodes;
	(function (ErrorCodes) {
		ErrorCodes.generalException="GeneralException";
	})(ErrorCodes=OfficeCore.ErrorCodes || (OfficeCore.ErrorCodes={}));
})(OfficeCore || (OfficeCore={}));
var OfficeCore;
(function (OfficeCore) {
	var _hostName="OfficeCore";
	var _defaultApiSetName="ExperimentApi";
	var _createPropertyObjectPath=OfficeExtension.ObjectPathFactory.createPropertyObjectPath;
	var _createMethodObjectPath=OfficeExtension.ObjectPathFactory.createMethodObjectPath;
	var _createIndexerObjectPath=OfficeExtension.ObjectPathFactory.createIndexerObjectPath;
	var _createNewObjectObjectPath=OfficeExtension.ObjectPathFactory.createNewObjectObjectPath;
	var _createChildItemObjectPathUsingIndexer=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexer;
	var _createChildItemObjectPathUsingGetItemAt=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingGetItemAt;
	var _createChildItemObjectPathUsingIndexerOrGetItemAt=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt;
	var _createMethodAction=OfficeExtension.ActionFactory.createMethodAction;
	var _createSetPropertyAction=OfficeExtension.ActionFactory.createSetPropertyAction;
	var _isNullOrUndefined=OfficeExtension.Utility.isNullOrUndefined;
	var _isUndefined=OfficeExtension.Utility.isUndefined;
	var _throwIfNotLoaded=OfficeExtension.Utility.throwIfNotLoaded;
	var _throwIfApiNotSupported=OfficeExtension.Utility.throwIfApiNotSupported;
	var _load=OfficeExtension.Utility.load;
	var _fixObjectPathIfNecessary=OfficeExtension.Utility.fixObjectPathIfNecessary;
	var _addActionResultHandler=OfficeExtension.Utility._addActionResultHandler;
	var _handleNavigationPropertyResults=OfficeExtension.Utility._handleNavigationPropertyResults;
	var _adjustToDateTime=OfficeExtension.Utility.adjustToDateTime;
	var FlightingService=(function (_super) {
		__extends(FlightingService, _super);
		function FlightingService() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(FlightingService.prototype, "_className", {
			get: function () {
				return "FlightingService";
			},
			enumerable: true,
			configurable: true
		});
		FlightingService.prototype.getClientSessionId=function () {
			var action=_createMethodAction(this.context, this, "GetClientSessionId", 1, []);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		FlightingService.prototype.getDeferredFlights=function () {
			var action=_createMethodAction(this.context, this, "GetDeferredFlights", 1, []);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		FlightingService.prototype.getFeature=function (featureName, type, defaultValue, possibleValues) {
			return new OfficeCore.ABType(this.context, _createMethodObjectPath(this.context, this, "GetFeature", 1, [featureName, type, defaultValue, possibleValues], false, false, null));
		};
		FlightingService.prototype.getFeatureGate=function (featureName, scope) {
			return new OfficeCore.ABType(this.context, _createMethodObjectPath(this.context, this, "GetFeatureGate", 1, [featureName, scope], false, false, null));
		};
		FlightingService.prototype.resetOverride=function (featureName) {
			_createMethodAction(this.context, this, "ResetOverride", 0, [featureName]);
		};
		FlightingService.prototype.setOverride=function (featureName, type, value) {
			_createMethodAction(this.context, this, "SetOverride", 0, [featureName, type, value]);
		};
		FlightingService.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
		};
		FlightingService.newObject=function (context) {
			var ret=new OfficeCore.FlightingService(context, _createNewObjectObjectPath(context, "Microsoft.Experiment.FlightingService", false));
			return ret;
		};
		FlightingService.prototype.toJSON=function () {
			return {};
		};
		return FlightingService;
	}(OfficeExtension.ClientObject));
	OfficeCore.FlightingService=FlightingService;
	var ABType=(function (_super) {
		__extends(ABType, _super);
		function ABType() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ABType.prototype, "_className", {
			get: function () {
				return "ABType";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ABType.prototype, "value", {
			get: function () {
				_throwIfNotLoaded("value", this.m_value, "ABType", this._isNull);
				return this.m_value;
			},
			enumerable: true,
			configurable: true
		});
		ABType.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Value"])) {
				this.m_value=obj["Value"];
			}
		};
		ABType.prototype.load=function (option) {
			_load(this, option);
			return this;
		};
		ABType.prototype.toJSON=function () {
			return {
				"value": this.m_value
			};
		};
		return ABType;
	}(OfficeExtension.ClientObject));
	OfficeCore.ABType=ABType;
	var FeatureType;
	(function (FeatureType) {
		FeatureType.boolean="Boolean";
		FeatureType.integer="Integer";
		FeatureType.string="String";
	})(FeatureType=OfficeCore.FeatureType || (OfficeCore.FeatureType={}));
	var ExperimentErrorCodes;
	(function (ExperimentErrorCodes) {
		ExperimentErrorCodes.generalException="GeneralException";
	})(ExperimentErrorCodes=OfficeCore.ExperimentErrorCodes || (OfficeCore.ExperimentErrorCodes={}));
})(OfficeCore || (OfficeCore={}));
var OfficeCore;
(function (OfficeCore) {
	var FirstPartyApis=(function () {
		function FirstPartyApis(context) {
			this.context=context;
		}
		Object.defineProperty(FirstPartyApis.prototype, "authentication", {
			get: function () {
				if (!this.m_authentication) {
					this.m_authentication=OfficeCore.AuthenticationService.newObject(this.context);
				}
				return this.m_authentication;
			},
			enumerable: true,
			configurable: true
		});
		return FirstPartyApis;
	}());
	OfficeCore.FirstPartyApis=FirstPartyApis;
	var RequestContext=(function (_super) {
		__extends(RequestContext, _super);
		function RequestContext(url) {
			return _super.call(this, url) || this;
		}
		Object.defineProperty(RequestContext.prototype, "firstParty", {
			get: function () {
				if (!this.m_firstPartyApis) {
					this.m_firstPartyApis=new FirstPartyApis(this);
				}
				return this.m_firstPartyApis;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RequestContext.prototype, "flighting", {
			get: function () {
				return this.flightingService;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RequestContext.prototype, "telemetry", {
			get: function () {
				if (!this.m_telemetry) {
					this.m_telemetry=OfficeCore.TelemetryService.newObject(this);
				}
				return this.m_telemetry;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RequestContext.prototype, "bi", {
			get: function () {
				if (!this.m_biShim) {
					this.m_biShim=OfficeCore.BiShim.newObject(this);
				}
				return this.m_biShim;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RequestContext.prototype, "flightingService", {
			get: function () {
				if (!this.m_flightingService) {
					this.m_flightingService=OfficeCore.FlightingService.newObject(this);
				}
				return this.m_flightingService;
			},
			enumerable: true,
			configurable: true
		});
		return RequestContext;
	}(OfficeExtension.ClientRequestContext));
	OfficeCore.RequestContext=RequestContext;
	function run(arg1, arg2) {
		return OfficeExtension.ClientRequestContext._runBatch("OfficeCore.run", arguments, function (requestInfo) { return new OfficeCore.RequestContext(requestInfo); });
	}
	OfficeCore.run=run;
})(OfficeCore || (OfficeCore={}));
var OfficeCore;
(function (OfficeCore) {
	var _hostName="OfficeCore";
	var _defaultApiSetName="TelemetryApi";
	var _createPropertyObjectPath=OfficeExtension.ObjectPathFactory.createPropertyObjectPath;
	var _createMethodObjectPath=OfficeExtension.ObjectPathFactory.createMethodObjectPath;
	var _createIndexerObjectPath=OfficeExtension.ObjectPathFactory.createIndexerObjectPath;
	var _createNewObjectObjectPath=OfficeExtension.ObjectPathFactory.createNewObjectObjectPath;
	var _createChildItemObjectPathUsingIndexer=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexer;
	var _createChildItemObjectPathUsingGetItemAt=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingGetItemAt;
	var _createChildItemObjectPathUsingIndexerOrGetItemAt=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt;
	var _createMethodAction=OfficeExtension.ActionFactory.createMethodAction;
	var _createSetPropertyAction=OfficeExtension.ActionFactory.createSetPropertyAction;
	var _isNullOrUndefined=OfficeExtension.Utility.isNullOrUndefined;
	var _isUndefined=OfficeExtension.Utility.isUndefined;
	var _throwIfNotLoaded=OfficeExtension.Utility.throwIfNotLoaded;
	var _throwIfApiNotSupported=OfficeExtension.Utility.throwIfApiNotSupported;
	var _load=OfficeExtension.Utility.load;
	var _fixObjectPathIfNecessary=OfficeExtension.Utility.fixObjectPathIfNecessary;
	var _addActionResultHandler=OfficeExtension.Utility._addActionResultHandler;
	var _handleNavigationPropertyResults=OfficeExtension.Utility._handleNavigationPropertyResults;
	var _adjustToDateTime=OfficeExtension.Utility.adjustToDateTime;
	var _typeTelemetryService="TelemetryService";
	var TelemetryService=(function (_super) {
		__extends(TelemetryService, _super);
		function TelemetryService() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(TelemetryService.prototype, "_className", {
			get: function () {
				return "TelemetryService";
			},
			enumerable: true,
			configurable: true
		});
		TelemetryService.prototype.sendTelemetryEvent=function (telemetryProperties, eventName, eventContract, eventFlags, value) {
			_createMethodAction(this.context, this, "SendTelemetryEvent", 1, [telemetryProperties, eventName, eventContract, eventFlags, value], false);
		};
		TelemetryService.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
		};
		TelemetryService.newObject=function (context) {
			var ret=new OfficeCore.TelemetryService(context, _createNewObjectObjectPath(context, "Microsoft.Telemetry.TelemetryService", false, false));
			return ret;
		};
		TelemetryService.prototype.toJSON=function () {
			return {};
		};
		return TelemetryService;
	}(OfficeExtension.ClientObject));
	OfficeCore.TelemetryService=TelemetryService;
	var TelemetryErrorCodes;
	(function (TelemetryErrorCodes) {
		TelemetryErrorCodes.generalException="GeneralException";
	})(TelemetryErrorCodes=OfficeCore.TelemetryErrorCodes || (OfficeCore.TelemetryErrorCodes={}));
})(OfficeCore || (OfficeCore={}));
var OfficeCore;
(function (OfficeCore) {
	var _hostName="Office";
	var _defaultApiSetName="OfficeSharedApi";
	var _createPropertyObjectPath=OfficeExtension.ObjectPathFactory.createPropertyObjectPath;
	var _createMethodObjectPath=OfficeExtension.ObjectPathFactory.createMethodObjectPath;
	var _createIndexerObjectPath=OfficeExtension.ObjectPathFactory.createIndexerObjectPath;
	var _createNewObjectObjectPath=OfficeExtension.ObjectPathFactory.createNewObjectObjectPath;
	var _createChildItemObjectPathUsingIndexer=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexer;
	var _createChildItemObjectPathUsingGetItemAt=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingGetItemAt;
	var _createChildItemObjectPathUsingIndexerOrGetItemAt=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt;
	var _createMethodAction=OfficeExtension.ActionFactory.createMethodAction;
	var _createEnsureUnchangedAction=OfficeExtension.ActionFactory.createEnsureUnchangedAction;
	var _createSetPropertyAction=OfficeExtension.ActionFactory.createSetPropertyAction;
	var _isNullOrUndefined=OfficeExtension.Utility.isNullOrUndefined;
	var _isUndefined=OfficeExtension.Utility.isUndefined;
	var _throwIfNotLoaded=OfficeExtension.Utility.throwIfNotLoaded;
	var _throwIfApiNotSupported=OfficeExtension.Utility.throwIfApiNotSupported;
	var _load=OfficeExtension.Utility.load;
	var _retrieve=OfficeExtension.Utility.retrieve;
	var _toJson=OfficeExtension.Utility.toJson;
	var _fixObjectPathIfNecessary=OfficeExtension.Utility.fixObjectPathIfNecessary;
	var _addActionResultHandler=OfficeExtension.Utility._addActionResultHandler;
	var _handleNavigationPropertyResults=OfficeExtension.Utility._handleNavigationPropertyResults;
	var _adjustToDateTime=OfficeExtension.Utility.adjustToDateTime;
	var _processRetrieveResult=OfficeExtension.Utility.processRetrieveResult;
	var IdentityType;
	(function (IdentityType) {
		IdentityType.organizationAccount="OrganizationAccount";
		IdentityType.microsoftAccount="MicrosoftAccount";
	})(IdentityType=OfficeCore.IdentityType || (OfficeCore.IdentityType={}));
	var _typeAuthenticationService="AuthenticationService";
	var AuthenticationService=(function (_super) {
		__extends(AuthenticationService, _super);
		function AuthenticationService() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(AuthenticationService.prototype, "_className", {
			get: function () {
				return "AuthenticationService";
			},
			enumerable: true,
			configurable: true
		});
		AuthenticationService.prototype.getAccessToken=function (tokenParameters) {
			var action=_createMethodAction(this.context, this, "GetAccessToken", 1, [tokenParameters], true);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		AuthenticationService.prototype.getPrimaryIdentityType=function () {
			var action=_createMethodAction(this.context, this, "GetPrimaryIdentityType", 1, [], true);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		AuthenticationService.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
		};
		AuthenticationService.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result);
		};
		AuthenticationService.newObject=function (context) {
			var ret=new OfficeCore.AuthenticationService(context, _createNewObjectObjectPath(context, "Microsoft.Authentication.AuthenticationService", false, false));
			return ret;
		};
		AuthenticationService.prototype.toJSON=function () {
			return _toJson(this, {}, {});
		};
		return AuthenticationService;
	}(OfficeExtension.ClientObject));
	OfficeCore.AuthenticationService=AuthenticationService;
	var _typeComment="Comment";
	var Comment=(function (_super) {
		__extends(Comment, _super);
		function Comment() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Comment.prototype, "_className", {
			get: function () {
				return "Comment";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["id", "text", "created", "level", "resolved", "author", "mentions"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, true, false, false, true, false, false];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["parent", "parentOrNullObject", "replies"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "parent", {
			get: function () {
				if (!this._P) {
					this._P=new OfficeCore.Comment(this.context, _createPropertyObjectPath(this.context, this, "Parent", false, false, false));
				}
				return this._P;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "parentOrNullObject", {
			get: function () {
				if (!this._Pa) {
					this._Pa=new OfficeCore.Comment(this.context, _createPropertyObjectPath(this.context, this, "ParentOrNullObject", false, false, false));
				}
				return this._Pa;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "replies", {
			get: function () {
				if (!this._R) {
					this._R=new OfficeCore.CommentCollection(this.context, _createPropertyObjectPath(this.context, this, "Replies", true, false, false));
				}
				return this._R;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "author", {
			get: function () {
				_throwIfNotLoaded("author", this._A, _typeComment, this._isNull);
				return this._A;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "created", {
			get: function () {
				_throwIfNotLoaded("created", this._C, _typeComment, this._isNull);
				return this._C;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "id", {
			get: function () {
				_throwIfNotLoaded("id", this._I, _typeComment, this._isNull);
				return this._I;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "level", {
			get: function () {
				_throwIfNotLoaded("level", this._L, _typeComment, this._isNull);
				return this._L;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "mentions", {
			get: function () {
				_throwIfNotLoaded("mentions", this._M, _typeComment, this._isNull);
				return this._M;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "resolved", {
			get: function () {
				_throwIfNotLoaded("resolved", this._Re, _typeComment, this._isNull);
				return this._Re;
			},
			set: function (value) {
				this._Re=value;
				_createSetPropertyAction(this.context, this, "Resolved", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Comment.prototype, "text", {
			get: function () {
				_throwIfNotLoaded("text", this._T, _typeComment, this._isNull);
				return this._T;
			},
			set: function (value) {
				this._T=value;
				_createSetPropertyAction(this.context, this, "Text", value);
			},
			enumerable: true,
			configurable: true
		});
		Comment.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["text", "resolved"], [], [
				"parent",
				"parentOrNullObject",
				"replies"
			]);
		};
		Comment.prototype.update=function (properties) {
			this._recursivelyUpdate(properties);
		};
		Comment.prototype.delete=function () {
			_createMethodAction(this.context, this, "Delete", 0, [], false);
		};
		Comment.prototype.getParentOrSelf=function () {
			return new OfficeCore.Comment(this.context, _createMethodObjectPath(this.context, this, "GetParentOrSelf", 1, [], false, false, null, false));
		};
		Comment.prototype.getRichText=function (format) {
			var action=_createMethodAction(this.context, this, "GetRichText", 1, [format], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		Comment.prototype.reply=function (text, format) {
			return new OfficeCore.Comment(this.context, _createMethodObjectPath(this.context, this, "Reply", 0, [text, format], false, false, null, false));
		};
		Comment.prototype.setRichText=function (text, format) {
			var action=_createMethodAction(this.context, this, "SetRichText", 0, [text, format], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		Comment.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Author"])) {
				this._A=obj["Author"];
			}
			if (!_isUndefined(obj["Created"])) {
				this._C=_adjustToDateTime(obj["Created"]);
			}
			if (!_isUndefined(obj["Id"])) {
				this._I=obj["Id"];
			}
			if (!_isUndefined(obj["Level"])) {
				this._L=obj["Level"];
			}
			if (!_isUndefined(obj["Mentions"])) {
				this._M=obj["Mentions"];
			}
			if (!_isUndefined(obj["Resolved"])) {
				this._Re=obj["Resolved"];
			}
			if (!_isUndefined(obj["Text"])) {
				this._T=obj["Text"];
			}
			_handleNavigationPropertyResults(this, obj, ["parent", "Parent", "parentOrNullObject", "ParentOrNullObject", "replies", "Replies"]);
		};
		Comment.prototype.load=function (option) {
			return _load(this, option);
		};
		Comment.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		Comment.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["Id"])) {
				this._I=value["Id"];
			}
		};
		Comment.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			if (!_isUndefined(obj["Created"])) {
				obj["created"]=_adjustToDateTime(obj["created"]);
			}
			_processRetrieveResult(this, value, result);
		};
		Comment.prototype.toJSON=function () {
			return _toJson(this, {
				"author": this._A,
				"created": this._C,
				"id": this._I,
				"level": this._L,
				"mentions": this._M,
				"resolved": this._Re,
				"text": this._T,
			}, {
				"replies": this._R,
			});
		};
		Comment.prototype.ensureUnchanged=function (data) {
			_createEnsureUnchangedAction(this.context, this, data);
			return;
		};
		return Comment;
	}(OfficeExtension.ClientObject));
	OfficeCore.Comment=Comment;
	var _typeCommentCollection="CommentCollection";
	var CommentCollection=(function (_super) {
		__extends(CommentCollection, _super);
		function CommentCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(CommentCollection.prototype, "_className", {
			get: function () {
				return "CommentCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CommentCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CommentCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, _typeCommentCollection, this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		CommentCollection.prototype.getCount=function () {
			var action=_createMethodAction(this.context, this, "GetCount", 1, [], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		CommentCollection.prototype.getItem=function (id) {
			return new OfficeCore.Comment(this.context, _createIndexerObjectPath(this.context, this, [id]));
		};
		CommentCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new OfficeCore.Comment(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		CommentCollection.prototype.load=function (option) {
			return _load(this, option);
		};
		CommentCollection.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		CommentCollection.prototype._handleRetrieveResult=function (value, result) {
			var _this=this;
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result, function (childItemData, index) { return new OfficeCore.Comment(_this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, _this.context, _this, childItemData, index)); });
		};
		CommentCollection.prototype.toJSON=function () {
			return _toJson(this, {}, {}, this.m__items);
		};
		return CommentCollection;
	}(OfficeExtension.ClientObject));
	OfficeCore.CommentCollection=CommentCollection;
	var CommentTextFormat;
	(function (CommentTextFormat) {
		CommentTextFormat.plain="Plain";
		CommentTextFormat.markdown="Markdown";
		CommentTextFormat.delta="Delta";
	})(CommentTextFormat=OfficeCore.CommentTextFormat || (OfficeCore.CommentTextFormat={}));
	var ErrorCodes;
	(function (ErrorCodes) {
		ErrorCodes.apiNotAvailable="ApiNotAvailable";
		ErrorCodes.clientError="ClientError";
		ErrorCodes.generalException="GeneralException";
		ErrorCodes.invalidArgument="InvalidArgument";
		ErrorCodes.invalidGrant="InvalidGrant";
		ErrorCodes.invalidResourceUrl="InvalidResourceUrl";
		ErrorCodes.serverError="ServerError";
		ErrorCodes.unsupportedUserIdentity="UnsupportedUserIdentity";
		ErrorCodes.userNotSignedIn="UserNotSignedIn";
	})(ErrorCodes=OfficeCore.ErrorCodes || (OfficeCore.ErrorCodes={}));
})(OfficeCore || (OfficeCore={}));

var __extends=(this && this.__extends) || (function () {
	var extendStatics=Object.setPrototypeOf ||
		({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__=b; }) ||
		function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p]=b[p]; };
	return function (d, b) {
		extendStatics(d, b);
		function __() { this.constructor=d; }
		d.prototype=b===null ? Object.create(b) : (__.prototype=b.prototype, new __());
	};
})();
var Word;
(function (Word) {
	function _normalizeSearchOptions(context, searchOptions) {
		if (OfficeExtension.Utility.isNullOrUndefined(searchOptions)) {
			return null;
		}
		if (typeof (searchOptions) !="object") {
			OfficeExtension.Utility.throwError(OfficeExtension.ResourceStrings.invalidArgument, "searchOptions");
		}
		if (searchOptions instanceof Word.SearchOptions) {
			return searchOptions;
		}
		var newSearchOptions=Word.SearchOptions.newObject(context);
		for (var property in searchOptions) {
			if (searchOptions.hasOwnProperty(property)) {
				newSearchOptions[property]=searchOptions[property];
			}
		}
		return newSearchOptions;
	}
	var _hostName="Word";
	var _defaultApiSetName="WordApi";
	var _createPropertyObjectPath=OfficeExtension.ObjectPathFactory.createPropertyObjectPath;
	var _createMethodObjectPath=OfficeExtension.ObjectPathFactory.createMethodObjectPath;
	var _createIndexerObjectPath=OfficeExtension.ObjectPathFactory.createIndexerObjectPath;
	var _createNewObjectObjectPath=OfficeExtension.ObjectPathFactory.createNewObjectObjectPath;
	var _createChildItemObjectPathUsingIndexer=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexer;
	var _createChildItemObjectPathUsingGetItemAt=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingGetItemAt;
	var _createChildItemObjectPathUsingIndexerOrGetItemAt=OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt;
	var _createMethodAction=OfficeExtension.ActionFactory.createMethodAction;
	var _createEnsureUnchangedAction=OfficeExtension.ActionFactory.createEnsureUnchangedAction;
	var _createSetPropertyAction=OfficeExtension.ActionFactory.createSetPropertyAction;
	var _isNullOrUndefined=OfficeExtension.Utility.isNullOrUndefined;
	var _isUndefined=OfficeExtension.Utility.isUndefined;
	var _throwIfNotLoaded=OfficeExtension.Utility.throwIfNotLoaded;
	var _throwIfApiNotSupported=OfficeExtension.Utility.throwIfApiNotSupported;
	var _load=OfficeExtension.Utility.load;
	var _retrieve=OfficeExtension.Utility.retrieve;
	var _toJson=OfficeExtension.Utility.toJson;
	var _fixObjectPathIfNecessary=OfficeExtension.Utility.fixObjectPathIfNecessary;
	var _addActionResultHandler=OfficeExtension.Utility._addActionResultHandler;
	var _handleNavigationPropertyResults=OfficeExtension.Utility._handleNavigationPropertyResults;
	var _adjustToDateTime=OfficeExtension.Utility.adjustToDateTime;
	var _processRetrieveResult=OfficeExtension.Utility.processRetrieveResult;
	var _typeApplication="Application";
	var Application=(function (_super) {
		__extends(Application, _super);
		function Application() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Application.prototype, "_className", {
			get: function () {
				return "Application";
			},
			enumerable: true,
			configurable: true
		});
		Application.prototype.createDocument=function (base64File) {
			return new Word.DocumentCreated(this.context, _createMethodObjectPath(this.context, this, "CreateDocument", 1, [base64File], false, false, null, false));
		};
		Application.prototype.getSharePointTenantRoot=function () {
			var action=_createMethodAction(this.context, this, "GetSharePointTenantRoot", 1, [], true);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		Application.prototype.isTapEnabled=function () {
			var action=_createMethodAction(this.context, this, "IsTapEnabled", 1, [], true);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		Application.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
		};
		Application.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result);
		};
		Application.newObject=function (context) {
			var ret=new Word.Application(context, _createNewObjectObjectPath(context, "Microsoft.WordServices.Application", false, false));
			return ret;
		};
		Application.prototype.toJSON=function () {
			return _toJson(this, {}, {});
		};
		return Application;
	}(OfficeExtension.ClientObject));
	Word.Application=Application;
	var _typeBody="Body";
	var Body=(function (_super) {
		__extends(Body, _super);
		function Body() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Body.prototype, "_className", {
			get: function () {
				return "Body";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Body.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["_ReferenceId", "style", "text", "type", "styleBuiltIn"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Body.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, true, false, false, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Body.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["paragraphs", "contentControls", "parentContentControl", "font", "inlinePictures", "parentBody", "lists", "tables", "parentSection", "parentContentControlOrNullObject", "parentBodyOrNullObject", "parentSectionOrNullObject"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Body.prototype, "contentControls", {
			get: function () {
				if (!this._C) {
					this._C=new Word.ContentControlCollection(this.context, _createPropertyObjectPath(this.context, this, "ContentControls", true, false, false));
				}
				return this._C;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Body.prototype, "font", {
			get: function () {
				if (!this._F) {
					this._F=new Word.Font(this.context, _createPropertyObjectPath(this.context, this, "Font", false, false, false));
				}
				return this._F;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Body.prototype, "inlinePictures", {
			get: function () {
				if (!this._I) {
					this._I=new Word.InlinePictureCollection(this.context, _createPropertyObjectPath(this.context, this, "InlinePictures", true, false, false));
				}
				return this._I;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Body.prototype, "lists", {
			get: function () {
				_throwIfApiNotSupported("Body.lists", _defaultApiSetName, "1.3", _hostName);
				if (!this._L) {
					this._L=new Word.ListCollection(this.context, _createPropertyObjectPath(this.context, this, "Lists", true, false, false));
				}
				return this._L;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Body.prototype, "paragraphs", {
			get: function () {
				if (!this._P) {
					this._P=new Word.ParagraphCollection(this.context, _createPropertyObjectPath(this.context, this, "Paragraphs", true, false, false));
				}
				return this._P;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Body.prototype, "parentBody", {
			get: function () {
				_throwIfApiNotSupported("Body.parentBody", _defaultApiSetName, "1.3", _hostName);
				if (!this._Pa) {
					this._Pa=new Word.Body(this.context, _createPropertyObjectPath(this.context, this, "ParentBody", false, false, false));
				}
				return this._Pa;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Body.prototype, "parentBodyOrNullObject", {
			get: function () {
				_throwIfApiNotSupported("Body.parentBodyOrNullObject", _defaultApiSetName, "1.3", _hostName);
				if (!this._Par) {
					this._Par=new Word.Body(this.context, _createPropertyObjectPath(this.context, this, "ParentBodyOrNullObject", false, false, false));
				}
				return this._Par;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Body.prototype, "parentContentControl", {
			get: function () {
				if (!this._Pare) {
					this._Pare=new Word.ContentControl(this.context, _createPropertyObjectPath(this.context, this, "ParentContentControl", false, false, false));
				}
				return this._Pare;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Body.prototype, "parentContentControlOrNullObject", {
			get: function () {
				_throwIfApiNotSupported("Body.parentContentControlOrNullObject", _defaultApiSetName, "1.3", _hostName);
				if (!this._Paren) {
					this._Paren=new Word.ContentControl(this.context, _createPropertyObjectPath(this.context, this, "ParentContentControlOrNullObject", false, false, false));
				}
				return this._Paren;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Body.prototype, "parentSection", {
			get: function () {
				_throwIfApiNotSupported("Body.parentSection", _defaultApiSetName, "1.3", _hostName);
				if (!this._Parent) {
					this._Parent=new Word.Section(this.context, _createPropertyObjectPath(this.context, this, "ParentSection", false, false, false));
				}
				return this._Parent;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Body.prototype, "parentSectionOrNullObject", {
			get: function () {
				_throwIfApiNotSupported("Body.parentSectionOrNullObject", _defaultApiSetName, "1.3", _hostName);
				if (!this._ParentS) {
					this._ParentS=new Word.Section(this.context, _createPropertyObjectPath(this.context, this, "ParentSectionOrNullObject", false, false, false));
				}
				return this._ParentS;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Body.prototype, "tables", {
			get: function () {
				_throwIfApiNotSupported("Body.tables", _defaultApiSetName, "1.3", _hostName);
				if (!this._T) {
					this._T=new Word.TableCollection(this.context, _createPropertyObjectPath(this.context, this, "Tables", true, false, false));
				}
				return this._T;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Body.prototype, "style", {
			get: function () {
				_throwIfNotLoaded("style", this._S, _typeBody, this._isNull);
				return this._S;
			},
			set: function (value) {
				this._S=value;
				_createSetPropertyAction(this.context, this, "Style", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Body.prototype, "styleBuiltIn", {
			get: function () {
				_throwIfNotLoaded("styleBuiltIn", this._St, _typeBody, this._isNull);
				_throwIfApiNotSupported("Body.styleBuiltIn", _defaultApiSetName, "1.3", _hostName);
				return this._St;
			},
			set: function (value) {
				this._St=value;
				_createSetPropertyAction(this.context, this, "StyleBuiltIn", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Body.prototype, "text", {
			get: function () {
				_throwIfNotLoaded("text", this._Te, _typeBody, this._isNull);
				return this._Te;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Body.prototype, "type", {
			get: function () {
				_throwIfNotLoaded("type", this._Ty, _typeBody, this._isNull);
				_throwIfApiNotSupported("Body.type", _defaultApiSetName, "1.3", _hostName);
				return this._Ty;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Body.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.__R, _typeBody, this._isNull);
				return this.__R;
			},
			enumerable: true,
			configurable: true
		});
		Body.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["style", "styleBuiltIn"], ["font"], [
				"contentControls",
				"inlinePictures",
				"lists",
				"paragraphs",
				"parentBody",
				"parentBodyOrNullObject",
				"parentContentControl",
				"parentContentControlOrNullObject",
				"parentSection",
				"parentSectionOrNullObject",
				"tables"
			]);
		};
		Body.prototype.update=function (properties) {
			this._recursivelyUpdate(properties);
		};
		Body.prototype.clear=function () {
			_createMethodAction(this.context, this, "Clear", 0, [], false);
		};
		Body.prototype.getHtml=function () {
			var action=_createMethodAction(this.context, this, "GetHtml", 1, [], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		Body.prototype.getOoxml=function () {
			var action=_createMethodAction(this.context, this, "GetOoxml", 1, [], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		Body.prototype.getRange=function (rangeLocation) {
			_throwIfApiNotSupported("Body.getRange", _defaultApiSetName, "1.3", _hostName);
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "GetRange", 1, [rangeLocation], false, false, null, false));
		};
		Body.prototype.insertBreak=function (breakType, insertLocation) {
			_createMethodAction(this.context, this, "InsertBreak", 0, [breakType, insertLocation], false);
		};
		Body.prototype.insertContentControl=function () {
			return new Word.ContentControl(this.context, _createMethodObjectPath(this.context, this, "InsertContentControl", 0, [], false, true, null, false));
		};
		Body.prototype.insertFileFromBase64=function (base64File, insertLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertFileFromBase64", 0, [base64File, insertLocation], false, true, null, false));
		};
		Body.prototype.insertHtml=function (html, insertLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertHtml", 0, [html, insertLocation], false, true, null, false));
		};
		Body.prototype.insertInlinePictureFromBase64=function (base64EncodedImage, insertLocation) {
			_throwIfApiNotSupported("Body.insertInlinePictureFromBase64", _defaultApiSetName, "1.2", _hostName);
			return new Word.InlinePicture(this.context, _createMethodObjectPath(this.context, this, "InsertInlinePictureFromBase64", 0, [base64EncodedImage, insertLocation], false, true, null, false));
		};
		Body.prototype.insertOoxml=function (ooxml, insertLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertOoxml", 0, [ooxml, insertLocation], false, true, null, false));
		};
		Body.prototype.insertParagraph=function (paragraphText, insertLocation) {
			return new Word.Paragraph(this.context, _createMethodObjectPath(this.context, this, "InsertParagraph", 0, [paragraphText, insertLocation], false, true, null, false));
		};
		Body.prototype.insertTable=function (rowCount, columnCount, insertLocation, values) {
			_throwIfApiNotSupported("Body.insertTable", _defaultApiSetName, "1.3", _hostName);
			return new Word.Table(this.context, _createMethodObjectPath(this.context, this, "InsertTable", 0, [rowCount, columnCount, insertLocation, values], false, true, null, false));
		};
		Body.prototype.insertText=function (text, insertLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertText", 0, [text, insertLocation], false, true, null, false));
		};
		Body.prototype.search=function (searchText, searchOptions) {
			searchOptions=_normalizeSearchOptions(this.context, searchOptions);
			return new Word.RangeCollection(this.context, _createMethodObjectPath(this.context, this, "Search", 1, [searchText, searchOptions], true, false, null, false));
		};
		Body.prototype.select=function (selectionMode) {
			_createMethodAction(this.context, this, "Select", 1, [selectionMode], false);
		};
		Body.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1, [], false);
		};
		Body.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Style"])) {
				this._S=obj["Style"];
			}
			if (!_isUndefined(obj["StyleBuiltIn"])) {
				this._St=obj["StyleBuiltIn"];
			}
			if (!_isUndefined(obj["Text"])) {
				this._Te=obj["Text"];
			}
			if (!_isUndefined(obj["Type"])) {
				this._Ty=obj["Type"];
			}
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.__R=obj["_ReferenceId"];
			}
			_handleNavigationPropertyResults(this, obj, ["contentControls", "ContentControls", "font", "Font", "inlinePictures", "InlinePictures", "lists", "Lists", "paragraphs", "Paragraphs", "parentBody", "ParentBody", "parentBodyOrNullObject", "ParentBodyOrNullObject", "parentContentControl", "ParentContentControl", "parentContentControlOrNullObject", "ParentContentControlOrNullObject", "parentSection", "ParentSection", "parentSectionOrNullObject", "ParentSectionOrNullObject", "tables", "Tables"]);
		};
		Body.prototype.load=function (option) {
			return _load(this, option);
		};
		Body.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		Body.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["_ReferenceId"])) {
				this.__R=value["_ReferenceId"];
			}
		};
		Body.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result);
		};
		Body.prototype.track=function () {
			this.context.trackedObjects.add(this);
			return this;
		};
		Body.prototype.untrack=function () {
			this.context.trackedObjects.remove(this);
			return this;
		};
		Body.prototype.toJSON=function () {
			return _toJson(this, {
				"style": this._S,
				"styleBuiltIn": this._St,
				"text": this._Te,
				"type": this._Ty,
			}, {
				"contentControls": this._C,
				"font": this._F,
				"inlinePictures": this._I,
				"lists": this._L,
				"paragraphs": this._P,
				"tables": this._T,
			});
		};
		Body.prototype.ensureUnchanged=function (data) {
			_createEnsureUnchangedAction(this.context, this, data);
			return;
		};
		return Body;
	}(OfficeExtension.ClientObject));
	Word.Body=Body;
	var _typeContentControl="ContentControl";
	var ContentControl=(function (_super) {
		__extends(ContentControl, _super);
		function ContentControl() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ContentControl.prototype, "_className", {
			get: function () {
				return "ContentControl";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["id", "_ReferenceId", "title", "tag", "placeholderText", "type", "appearance", "color", "removeWhenEdited", "cannotDelete", "cannotEdit", "style", "text", "subtype", "styleBuiltIn"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, false, true, true, true, false, true, true, true, true, true, true, false, false, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["font", "paragraphs", "contentControls", "parentContentControl", "inlinePictures", "lists", "tables", "parentTableCell", "parentTable", "parentBody", "parentContentControlOrNullObject", "parentTableCellOrNullObject", "parentTableOrNullObject"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "contentControls", {
			get: function () {
				if (!this._Con) {
					this._Con=new Word.ContentControlCollection(this.context, _createPropertyObjectPath(this.context, this, "ContentControls", true, false, false));
				}
				return this._Con;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "font", {
			get: function () {
				if (!this._F) {
					this._F=new Word.Font(this.context, _createPropertyObjectPath(this.context, this, "Font", false, false, false));
				}
				return this._F;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "inlinePictures", {
			get: function () {
				if (!this._In) {
					this._In=new Word.InlinePictureCollection(this.context, _createPropertyObjectPath(this.context, this, "InlinePictures", true, false, false));
				}
				return this._In;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "lists", {
			get: function () {
				_throwIfApiNotSupported("ContentControl.lists", _defaultApiSetName, "1.3", _hostName);
				if (!this._L) {
					this._L=new Word.ListCollection(this.context, _createPropertyObjectPath(this.context, this, "Lists", true, false, false));
				}
				return this._L;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "paragraphs", {
			get: function () {
				if (!this._P) {
					this._P=new Word.ParagraphCollection(this.context, _createPropertyObjectPath(this.context, this, "Paragraphs", true, false, false));
				}
				return this._P;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "parentBody", {
			get: function () {
				_throwIfApiNotSupported("ContentControl.parentBody", _defaultApiSetName, "1.3", _hostName);
				if (!this._Pa) {
					this._Pa=new Word.Body(this.context, _createPropertyObjectPath(this.context, this, "ParentBody", false, false, false));
				}
				return this._Pa;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "parentContentControl", {
			get: function () {
				if (!this._Par) {
					this._Par=new Word.ContentControl(this.context, _createPropertyObjectPath(this.context, this, "ParentContentControl", false, false, false));
				}
				return this._Par;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "parentContentControlOrNullObject", {
			get: function () {
				_throwIfApiNotSupported("ContentControl.parentContentControlOrNullObject", _defaultApiSetName, "1.3", _hostName);
				if (!this._Pare) {
					this._Pare=new Word.ContentControl(this.context, _createPropertyObjectPath(this.context, this, "ParentContentControlOrNullObject", false, false, false));
				}
				return this._Pare;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "parentTable", {
			get: function () {
				_throwIfApiNotSupported("ContentControl.parentTable", _defaultApiSetName, "1.3", _hostName);
				if (!this._Paren) {
					this._Paren=new Word.Table(this.context, _createPropertyObjectPath(this.context, this, "ParentTable", false, false, false));
				}
				return this._Paren;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "parentTableCell", {
			get: function () {
				_throwIfApiNotSupported("ContentControl.parentTableCell", _defaultApiSetName, "1.3", _hostName);
				if (!this._Parent) {
					this._Parent=new Word.TableCell(this.context, _createPropertyObjectPath(this.context, this, "ParentTableCell", false, false, false));
				}
				return this._Parent;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "parentTableCellOrNullObject", {
			get: function () {
				_throwIfApiNotSupported("ContentControl.parentTableCellOrNullObject", _defaultApiSetName, "1.3", _hostName);
				if (!this._ParentT) {
					this._ParentT=new Word.TableCell(this.context, _createPropertyObjectPath(this.context, this, "ParentTableCellOrNullObject", false, false, false));
				}
				return this._ParentT;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "parentTableOrNullObject", {
			get: function () {
				_throwIfApiNotSupported("ContentControl.parentTableOrNullObject", _defaultApiSetName, "1.3", _hostName);
				if (!this._ParentTa) {
					this._ParentTa=new Word.Table(this.context, _createPropertyObjectPath(this.context, this, "ParentTableOrNullObject", false, false, false));
				}
				return this._ParentTa;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "tables", {
			get: function () {
				_throwIfApiNotSupported("ContentControl.tables", _defaultApiSetName, "1.3", _hostName);
				if (!this._T) {
					this._T=new Word.TableCollection(this.context, _createPropertyObjectPath(this.context, this, "Tables", true, false, false));
				}
				return this._T;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "appearance", {
			get: function () {
				_throwIfNotLoaded("appearance", this._A, _typeContentControl, this._isNull);
				return this._A;
			},
			set: function (value) {
				this._A=value;
				_createSetPropertyAction(this.context, this, "Appearance", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "cannotDelete", {
			get: function () {
				_throwIfNotLoaded("cannotDelete", this._C, _typeContentControl, this._isNull);
				return this._C;
			},
			set: function (value) {
				this._C=value;
				_createSetPropertyAction(this.context, this, "CannotDelete", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "cannotEdit", {
			get: function () {
				_throwIfNotLoaded("cannotEdit", this._Ca, _typeContentControl, this._isNull);
				return this._Ca;
			},
			set: function (value) {
				this._Ca=value;
				_createSetPropertyAction(this.context, this, "CannotEdit", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "color", {
			get: function () {
				_throwIfNotLoaded("color", this._Co, _typeContentControl, this._isNull);
				return this._Co;
			},
			set: function (value) {
				this._Co=value;
				_createSetPropertyAction(this.context, this, "Color", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "id", {
			get: function () {
				_throwIfNotLoaded("id", this._I, _typeContentControl, this._isNull);
				return this._I;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "placeholderText", {
			get: function () {
				_throwIfNotLoaded("placeholderText", this._Pl, _typeContentControl, this._isNull);
				return this._Pl;
			},
			set: function (value) {
				this._Pl=value;
				_createSetPropertyAction(this.context, this, "PlaceholderText", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "removeWhenEdited", {
			get: function () {
				_throwIfNotLoaded("removeWhenEdited", this._R, _typeContentControl, this._isNull);
				return this._R;
			},
			set: function (value) {
				this._R=value;
				_createSetPropertyAction(this.context, this, "RemoveWhenEdited", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "style", {
			get: function () {
				_throwIfNotLoaded("style", this._S, _typeContentControl, this._isNull);
				return this._S;
			},
			set: function (value) {
				this._S=value;
				_createSetPropertyAction(this.context, this, "Style", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "styleBuiltIn", {
			get: function () {
				_throwIfNotLoaded("styleBuiltIn", this._St, _typeContentControl, this._isNull);
				_throwIfApiNotSupported("ContentControl.styleBuiltIn", _defaultApiSetName, "1.3", _hostName);
				return this._St;
			},
			set: function (value) {
				this._St=value;
				_createSetPropertyAction(this.context, this, "StyleBuiltIn", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "subtype", {
			get: function () {
				_throwIfNotLoaded("subtype", this._Su, _typeContentControl, this._isNull);
				_throwIfApiNotSupported("ContentControl.subtype", _defaultApiSetName, "1.3", _hostName);
				return this._Su;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "tag", {
			get: function () {
				_throwIfNotLoaded("tag", this._Ta, _typeContentControl, this._isNull);
				return this._Ta;
			},
			set: function (value) {
				this._Ta=value;
				_createSetPropertyAction(this.context, this, "Tag", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "text", {
			get: function () {
				_throwIfNotLoaded("text", this._Te, _typeContentControl, this._isNull);
				return this._Te;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "title", {
			get: function () {
				_throwIfNotLoaded("title", this._Ti, _typeContentControl, this._isNull);
				return this._Ti;
			},
			set: function (value) {
				this._Ti=value;
				_createSetPropertyAction(this.context, this, "Title", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "type", {
			get: function () {
				_throwIfNotLoaded("type", this._Ty, _typeContentControl, this._isNull);
				return this._Ty;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControl.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.__R, _typeContentControl, this._isNull);
				return this.__R;
			},
			enumerable: true,
			configurable: true
		});
		ContentControl.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["title", "tag", "placeholderText", "appearance", "color", "removeWhenEdited", "cannotDelete", "cannotEdit", "style", "styleBuiltIn"], ["font"], [
				"contentControls",
				"inlinePictures",
				"lists",
				"paragraphs",
				"parentBody",
				"parentContentControl",
				"parentContentControlOrNullObject",
				"parentTable",
				"parentTableCell",
				"parentTableCellOrNullObject",
				"parentTableOrNullObject",
				"tables"
			]);
		};
		ContentControl.prototype.update=function (properties) {
			this._recursivelyUpdate(properties);
		};
		ContentControl.prototype.clear=function () {
			_createMethodAction(this.context, this, "Clear", 0, [], false);
		};
		ContentControl.prototype.delete=function (keepContent) {
			_createMethodAction(this.context, this, "Delete", 0, [keepContent], false);
		};
		ContentControl.prototype.getHtml=function () {
			var action=_createMethodAction(this.context, this, "GetHtml", 1, [], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		ContentControl.prototype.getOoxml=function () {
			var action=_createMethodAction(this.context, this, "GetOoxml", 1, [], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		ContentControl.prototype.getRange=function (rangeLocation) {
			_throwIfApiNotSupported("ContentControl.getRange", _defaultApiSetName, "1.3", _hostName);
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "GetRange", 1, [rangeLocation], false, false, null, false));
		};
		ContentControl.prototype.getTextRanges=function (endingMarks, trimSpacing) {
			_throwIfApiNotSupported("ContentControl.getTextRanges", _defaultApiSetName, "1.3", _hostName);
			return new Word.RangeCollection(this.context, _createMethodObjectPath(this.context, this, "GetTextRanges", 1, [endingMarks, trimSpacing], true, false, null, false));
		};
		ContentControl.prototype.insertBreak=function (breakType, insertLocation) {
			_createMethodAction(this.context, this, "InsertBreak", 0, [breakType, insertLocation], false);
		};
		ContentControl.prototype.insertFileFromBase64=function (base64File, insertLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertFileFromBase64", 0, [base64File, insertLocation], false, true, null, false));
		};
		ContentControl.prototype.insertHtml=function (html, insertLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertHtml", 0, [html, insertLocation], false, true, null, false));
		};
		ContentControl.prototype.insertInlinePictureFromBase64=function (base64EncodedImage, insertLocation) {
			_throwIfApiNotSupported("ContentControl.insertInlinePictureFromBase64", _defaultApiSetName, "1.2", _hostName);
			return new Word.InlinePicture(this.context, _createMethodObjectPath(this.context, this, "InsertInlinePictureFromBase64", 0, [base64EncodedImage, insertLocation], false, true, null, false));
		};
		ContentControl.prototype.insertOoxml=function (ooxml, insertLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertOoxml", 0, [ooxml, insertLocation], false, true, null, false));
		};
		ContentControl.prototype.insertParagraph=function (paragraphText, insertLocation) {
			return new Word.Paragraph(this.context, _createMethodObjectPath(this.context, this, "InsertParagraph", 0, [paragraphText, insertLocation], false, true, null, false));
		};
		ContentControl.prototype.insertTable=function (rowCount, columnCount, insertLocation, values) {
			_throwIfApiNotSupported("ContentControl.insertTable", _defaultApiSetName, "1.3", _hostName);
			return new Word.Table(this.context, _createMethodObjectPath(this.context, this, "InsertTable", 0, [rowCount, columnCount, insertLocation, values], false, true, null, false));
		};
		ContentControl.prototype.insertText=function (text, insertLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertText", 0, [text, insertLocation], false, true, null, false));
		};
		ContentControl.prototype.search=function (searchText, searchOptions) {
			searchOptions=_normalizeSearchOptions(this.context, searchOptions);
			return new Word.RangeCollection(this.context, _createMethodObjectPath(this.context, this, "Search", 1, [searchText, searchOptions], true, false, null, false));
		};
		ContentControl.prototype.select=function (selectionMode) {
			_createMethodAction(this.context, this, "Select", 1, [selectionMode], false);
		};
		ContentControl.prototype.split=function (delimiters, multiParagraphs, trimDelimiters, trimSpacing) {
			_throwIfApiNotSupported("ContentControl.split", _defaultApiSetName, "1.3", _hostName);
			return new Word.RangeCollection(this.context, _createMethodObjectPath(this.context, this, "Split", 1, [delimiters, multiParagraphs, trimDelimiters, trimSpacing], true, false, null, false));
		};
		ContentControl.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1, [], false);
		};
		ContentControl.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Appearance"])) {
				this._A=obj["Appearance"];
			}
			if (!_isUndefined(obj["CannotDelete"])) {
				this._C=obj["CannotDelete"];
			}
			if (!_isUndefined(obj["CannotEdit"])) {
				this._Ca=obj["CannotEdit"];
			}
			if (!_isUndefined(obj["Color"])) {
				this._Co=obj["Color"];
			}
			if (!_isUndefined(obj["Id"])) {
				this._I=obj["Id"];
			}
			if (!_isUndefined(obj["PlaceholderText"])) {
				this._Pl=obj["PlaceholderText"];
			}
			if (!_isUndefined(obj["RemoveWhenEdited"])) {
				this._R=obj["RemoveWhenEdited"];
			}
			if (!_isUndefined(obj["Style"])) {
				this._S=obj["Style"];
			}
			if (!_isUndefined(obj["StyleBuiltIn"])) {
				this._St=obj["StyleBuiltIn"];
			}
			if (!_isUndefined(obj["Subtype"])) {
				this._Su=obj["Subtype"];
			}
			if (!_isUndefined(obj["Tag"])) {
				this._Ta=obj["Tag"];
			}
			if (!_isUndefined(obj["Text"])) {
				this._Te=obj["Text"];
			}
			if (!_isUndefined(obj["Title"])) {
				this._Ti=obj["Title"];
			}
			if (!_isUndefined(obj["Type"])) {
				this._Ty=obj["Type"];
			}
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.__R=obj["_ReferenceId"];
			}
			_handleNavigationPropertyResults(this, obj, ["contentControls", "ContentControls", "font", "Font", "inlinePictures", "InlinePictures", "lists", "Lists", "paragraphs", "Paragraphs", "parentBody", "ParentBody", "parentContentControl", "ParentContentControl", "parentContentControlOrNullObject", "ParentContentControlOrNullObject", "parentTable", "ParentTable", "parentTableCell", "ParentTableCell", "parentTableCellOrNullObject", "ParentTableCellOrNullObject", "parentTableOrNullObject", "ParentTableOrNullObject", "tables", "Tables"]);
		};
		ContentControl.prototype.load=function (option) {
			return _load(this, option);
		};
		ContentControl.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		ContentControl.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["_ReferenceId"])) {
				this.__R=value["_ReferenceId"];
			}
			if (!_isUndefined(value["Id"])) {
				this._I=value["Id"];
			}
		};
		ContentControl.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result);
		};
		ContentControl.prototype.track=function () {
			this.context.trackedObjects.add(this);
			return this;
		};
		ContentControl.prototype.untrack=function () {
			this.context.trackedObjects.remove(this);
			return this;
		};
		ContentControl.prototype.toJSON=function () {
			return _toJson(this, {
				"appearance": this._A,
				"cannotDelete": this._C,
				"cannotEdit": this._Ca,
				"color": this._Co,
				"id": this._I,
				"placeholderText": this._Pl,
				"removeWhenEdited": this._R,
				"style": this._S,
				"styleBuiltIn": this._St,
				"subtype": this._Su,
				"tag": this._Ta,
				"text": this._Te,
				"title": this._Ti,
				"type": this._Ty,
			}, {
				"contentControls": this._Con,
				"font": this._F,
				"inlinePictures": this._In,
				"lists": this._L,
				"paragraphs": this._P,
				"tables": this._T,
			});
		};
		ContentControl.prototype.ensureUnchanged=function (data) {
			_createEnsureUnchangedAction(this.context, this, data);
			return;
		};
		return ContentControl;
	}(OfficeExtension.ClientObject));
	Word.ContentControl=ContentControl;
	var _typeContentControlCollection="ContentControlCollection";
	var ContentControlCollection=(function (_super) {
		__extends(ContentControlCollection, _super);
		function ContentControlCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ContentControlCollection.prototype, "_className", {
			get: function () {
				return "ContentControlCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControlCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControlCollection.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["_ReferenceId"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControlCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, _typeContentControlCollection, this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ContentControlCollection.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.__R, _typeContentControlCollection, this._isNull);
				return this.__R;
			},
			enumerable: true,
			configurable: true
		});
		ContentControlCollection.prototype.getById=function (id) {
			return new Word.ContentControl(this.context, _createMethodObjectPath(this.context, this, "GetById", 1, [id], false, false, null, false));
		};
		ContentControlCollection.prototype.getByIdOrNullObject=function (id) {
			_throwIfApiNotSupported("ContentControlCollection.getByIdOrNullObject", _defaultApiSetName, "1.3", _hostName);
			return new Word.ContentControl(this.context, _createMethodObjectPath(this.context, this, "GetByIdOrNullObject", 1, [id], false, false, null, false));
		};
		ContentControlCollection.prototype.getByTag=function (tag) {
			return new Word.ContentControlCollection(this.context, _createMethodObjectPath(this.context, this, "GetByTag", 1, [tag], true, false, null, false));
		};
		ContentControlCollection.prototype.getByTitle=function (title) {
			return new Word.ContentControlCollection(this.context, _createMethodObjectPath(this.context, this, "GetByTitle", 1, [title], true, false, null, false));
		};
		ContentControlCollection.prototype.getByTypes=function (types) {
			_throwIfApiNotSupported("ContentControlCollection.getByTypes", _defaultApiSetName, "1.3", _hostName);
			return new Word.ContentControlCollection(this.context, _createMethodObjectPath(this.context, this, "GetByTypes", 1, [types], true, false, null, false));
		};
		ContentControlCollection.prototype.getFirst=function () {
			_throwIfApiNotSupported("ContentControlCollection.getFirst", _defaultApiSetName, "1.3", _hostName);
			return new Word.ContentControl(this.context, _createMethodObjectPath(this.context, this, "GetFirst", 1, [], false, false, null, false));
		};
		ContentControlCollection.prototype.getFirstOrNullObject=function () {
			_throwIfApiNotSupported("ContentControlCollection.getFirstOrNullObject", _defaultApiSetName, "1.3", _hostName);
			return new Word.ContentControl(this.context, _createMethodObjectPath(this.context, this, "GetFirstOrNullObject", 1, [], false, false, null, false));
		};
		ContentControlCollection.prototype.getItem=function (index) {
			return new Word.ContentControl(this.context, _createIndexerObjectPath(this.context, this, [index]));
		};
		ContentControlCollection.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1, [], false);
		};
		ContentControlCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.__R=obj["_ReferenceId"];
			}
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Word.ContentControl(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		ContentControlCollection.prototype.load=function (option) {
			return _load(this, option);
		};
		ContentControlCollection.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		ContentControlCollection.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["_ReferenceId"])) {
				this.__R=value["_ReferenceId"];
			}
		};
		ContentControlCollection.prototype._handleRetrieveResult=function (value, result) {
			var _this=this;
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result, function (childItemData, index) { return new Word.ContentControl(_this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, _this.context, _this, childItemData, index)); });
		};
		ContentControlCollection.prototype.track=function () {
			this.context.trackedObjects.add(this);
			return this;
		};
		ContentControlCollection.prototype.untrack=function () {
			this.context.trackedObjects.remove(this);
			return this;
		};
		ContentControlCollection.prototype.toJSON=function () {
			return _toJson(this, {}, {}, this.m__items);
		};
		return ContentControlCollection;
	}(OfficeExtension.ClientObject));
	Word.ContentControlCollection=ContentControlCollection;
	var _typeCustomProperty="CustomProperty";
	var CustomProperty=(function (_super) {
		__extends(CustomProperty, _super);
		function CustomProperty() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(CustomProperty.prototype, "_className", {
			get: function () {
				return "CustomProperty";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CustomProperty.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["_ReferenceId", "key", "value", "type", "_Id"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CustomProperty.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, false, true, false, false];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CustomProperty.prototype, "key", {
			get: function () {
				_throwIfNotLoaded("key", this._K, _typeCustomProperty, this._isNull);
				return this._K;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CustomProperty.prototype, "type", {
			get: function () {
				_throwIfNotLoaded("type", this.m_type, _typeCustomProperty, this._isNull);
				return this.m_type;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CustomProperty.prototype, "value", {
			get: function () {
				_throwIfNotLoaded("type", this.m_type, "CustomProperty", this._isNull);
				if (this.m_type=="Date") {
					_throwIfNotLoaded("value", this.m_value, "CustomProperty", this._isNull);
					return new Date(this.m_value);
				}
				_throwIfNotLoaded("value", this.m_value, _typeCustomProperty, this._isNull);
				return this.m_value;
			},
			set: function (value) {
				this.m_value=value;
				_createSetPropertyAction(this.context, this, "Value", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CustomProperty.prototype, "_Id", {
			get: function () {
				_throwIfNotLoaded("_Id", this.__I, _typeCustomProperty, this._isNull);
				return this.__I;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CustomProperty.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.__R, _typeCustomProperty, this._isNull);
				return this.__R;
			},
			enumerable: true,
			configurable: true
		});
		CustomProperty.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["value"], [], []);
		};
		CustomProperty.prototype.update=function (properties) {
			this._recursivelyUpdate(properties);
		};
		CustomProperty.prototype.delete=function () {
			_createMethodAction(this.context, this, "Delete", 0, [], false);
		};
		CustomProperty.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1, [], false);
		};
		CustomProperty.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Key"])) {
				this._K=obj["Key"];
			}
			if (!_isUndefined(obj["Type"])) {
				this.m_type=obj["Type"];
			}
			if (!_isUndefined(obj["Value"])) {
				this.m_value=obj["Value"];
			}
			if (!_isUndefined(obj["_Id"])) {
				this.__I=obj["_Id"];
			}
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.__R=obj["_ReferenceId"];
			}
		};
		CustomProperty.prototype.load=function (option) {
			return _load(this, option);
		};
		CustomProperty.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		CustomProperty.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["_ReferenceId"])) {
				this.__R=value["_ReferenceId"];
			}
			if (!_isUndefined(value["_Id"])) {
				this.__I=value["_Id"];
			}
		};
		CustomProperty.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result);
		};
		CustomProperty.prototype.track=function () {
			this.context.trackedObjects.add(this);
			return this;
		};
		CustomProperty.prototype.untrack=function () {
			this.context.trackedObjects.remove(this);
			return this;
		};
		CustomProperty.prototype.toJSON=function () {
			return _toJson(this, {
				"key": this._K,
				"type": this.m_type,
				"value": this.m_value,
			}, {});
		};
		CustomProperty.prototype.ensureUnchanged=function (data) {
			_createEnsureUnchangedAction(this.context, this, data);
			return;
		};
		return CustomProperty;
	}(OfficeExtension.ClientObject));
	Word.CustomProperty=CustomProperty;
	var _typeCustomPropertyCollection="CustomPropertyCollection";
	var CustomPropertyCollection=(function (_super) {
		__extends(CustomPropertyCollection, _super);
		function CustomPropertyCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(CustomPropertyCollection.prototype, "_className", {
			get: function () {
				return "CustomPropertyCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CustomPropertyCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CustomPropertyCollection.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["_ReferenceId"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CustomPropertyCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, _typeCustomPropertyCollection, this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(CustomPropertyCollection.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.__R, _typeCustomPropertyCollection, this._isNull);
				return this.__R;
			},
			enumerable: true,
			configurable: true
		});
		CustomPropertyCollection.prototype.add=function (key, value) {
			return new Word.CustomProperty(this.context, _createMethodObjectPath(this.context, this, "Add", 0, [key, value], false, true, null, false));
		};
		CustomPropertyCollection.prototype.deleteAll=function () {
			_createMethodAction(this.context, this, "DeleteAll", 0, [], false);
		};
		CustomPropertyCollection.prototype.getCount=function () {
			var action=_createMethodAction(this.context, this, "GetCount", 1, [], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		CustomPropertyCollection.prototype.getItem=function (key) {
			return new Word.CustomProperty(this.context, _createIndexerObjectPath(this.context, this, [key]));
		};
		CustomPropertyCollection.prototype.getItemOrNullObject=function (key) {
			return new Word.CustomProperty(this.context, _createMethodObjectPath(this.context, this, "GetItemOrNullObject", 1, [key], false, false, null, false));
		};
		CustomPropertyCollection.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1, [], false);
		};
		CustomPropertyCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.__R=obj["_ReferenceId"];
			}
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Word.CustomProperty(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		CustomPropertyCollection.prototype.load=function (option) {
			return _load(this, option);
		};
		CustomPropertyCollection.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		CustomPropertyCollection.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["_ReferenceId"])) {
				this.__R=value["_ReferenceId"];
			}
		};
		CustomPropertyCollection.prototype._handleRetrieveResult=function (value, result) {
			var _this=this;
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result, function (childItemData, index) { return new Word.CustomProperty(_this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, _this.context, _this, childItemData, index)); });
		};
		CustomPropertyCollection.prototype.track=function () {
			this.context.trackedObjects.add(this);
			return this;
		};
		CustomPropertyCollection.prototype.untrack=function () {
			this.context.trackedObjects.remove(this);
			return this;
		};
		CustomPropertyCollection.prototype.toJSON=function () {
			return _toJson(this, {}, {}, this.m__items);
		};
		return CustomPropertyCollection;
	}(OfficeExtension.ClientObject));
	Word.CustomPropertyCollection=CustomPropertyCollection;
	var _typeDocument="Document";
	var Document=(function (_super) {
		__extends(Document, _super);
		function Document() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Document.prototype, "_className", {
			get: function () {
				return "Document";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Document.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["saved", "_ReferenceId", "allowCloseOnUntitled"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Document.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, false, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Document.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["sections", "body", "contentControls", "properties"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Document.prototype, "body", {
			get: function () {
				if (!this._B) {
					this._B=new Word.Body(this.context, _createPropertyObjectPath(this.context, this, "Body", false, false, false));
				}
				return this._B;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Document.prototype, "contentControls", {
			get: function () {
				if (!this._C) {
					this._C=new Word.ContentControlCollection(this.context, _createPropertyObjectPath(this.context, this, "ContentControls", true, false, false));
				}
				return this._C;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Document.prototype, "properties", {
			get: function () {
				_throwIfApiNotSupported("Document.properties", _defaultApiSetName, "1.3", _hostName);
				if (!this._P) {
					this._P=new Word.DocumentProperties(this.context, _createPropertyObjectPath(this.context, this, "Properties", false, false, false));
				}
				return this._P;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Document.prototype, "sections", {
			get: function () {
				if (!this._Se) {
					this._Se=new Word.SectionCollection(this.context, _createPropertyObjectPath(this.context, this, "Sections", true, false, false));
				}
				return this._Se;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Document.prototype, "allowCloseOnUntitled", {
			get: function () {
				_throwIfNotLoaded("allowCloseOnUntitled", this._A, _typeDocument, this._isNull);
				return this._A;
			},
			set: function (value) {
				this._A=value;
				_createSetPropertyAction(this.context, this, "AllowCloseOnUntitled", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Document.prototype, "saved", {
			get: function () {
				_throwIfNotLoaded("saved", this._S, _typeDocument, this._isNull);
				return this._S;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Document.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.m__ReferenceId, _typeDocument, this._isNull);
				return this.m__ReferenceId;
			},
			enumerable: true,
			configurable: true
		});
		Document.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["allowCloseOnUntitled"], ["body", "properties"], [
				"contentControls",
				"sections"
			]);
		};
		Document.prototype.update=function (properties) {
			this._recursivelyUpdate(properties);
		};
		Document.prototype.getMetadata=function (propertyGUID) {
			var action=_createMethodAction(this.context, this, "GetMetadata", 1, [propertyGUID], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		Document.prototype.getNeighborhoodTextAroundSelection=function (maxLength) {
			var action=_createMethodAction(this.context, this, "GetNeighborhoodTextAroundSelection", 1, [maxLength], true);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		Document.prototype.getSelection=function () {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "GetSelection", 1, [], false, true, null, false));
		};
		Document.prototype.launchTapPane=function (documentUrl) {
			_createMethodAction(this.context, this, "LaunchTapPane", 1, [documentUrl], true);
		};
		Document.prototype.save=function () {
			_createMethodAction(this.context, this, "Save", 0, [], false);
		};
		Document.prototype.setMetadata=function (propertyGUID, propertyValue) {
			_createMethodAction(this.context, this, "SetMetadata", 0, [propertyGUID, propertyValue], false);
		};
		Document.prototype.setMetadataOnTile=function (tileID, propertyGUID, propertyValue) {
			_createMethodAction(this.context, this, "SetMetadataOnTile", 0, [tileID, propertyGUID, propertyValue], false);
		};
		Document.prototype._GetObjectByReferenceId=function (referenceId) {
			var action=_createMethodAction(this.context, this, "_GetObjectByReferenceId", 1, [referenceId], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		Document.prototype._GetObjectTypeNameByReferenceId=function (referenceId) {
			var action=_createMethodAction(this.context, this, "_GetObjectTypeNameByReferenceId", 1, [referenceId], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		Document.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1, [], false);
		};
		Document.prototype._RemoveAllReferences=function () {
			_createMethodAction(this.context, this, "_RemoveAllReferences", 1, [], false);
		};
		Document.prototype._RemoveReference=function (referenceId) {
			_createMethodAction(this.context, this, "_RemoveReference", 1, [referenceId], false);
		};
		Document.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["AllowCloseOnUntitled"])) {
				this._A=obj["AllowCloseOnUntitled"];
			}
			if (!_isUndefined(obj["Saved"])) {
				this._S=obj["Saved"];
			}
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.m__ReferenceId=obj["_ReferenceId"];
			}
			_handleNavigationPropertyResults(this, obj, ["body", "Body", "contentControls", "ContentControls", "properties", "Properties", "sections", "Sections"]);
		};
		Document.prototype.load=function (option) {
			return _load(this, option);
		};
		Document.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		Document.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["_ReferenceId"])) {
				this.m__ReferenceId=value["_ReferenceId"];
			}
		};
		Document.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result);
		};
		Document.prototype.track=function () {
			this.context.trackedObjects.add(this);
			return this;
		};
		Document.prototype.untrack=function () {
			this.context.trackedObjects.remove(this);
			return this;
		};
		Document.prototype.toJSON=function () {
			return _toJson(this, {
				"allowCloseOnUntitled": this._A,
				"saved": this._S,
			}, {
				"body": this._B,
				"contentControls": this._C,
				"properties": this._P,
				"sections": this._Se,
			});
		};
		Document.prototype.ensureUnchanged=function (data) {
			_createEnsureUnchangedAction(this.context, this, data);
			return;
		};
		return Document;
	}(OfficeExtension.ClientObject));
	Word.Document=Document;
	var _typeDocumentCreated="DocumentCreated";
	var DocumentCreated=(function (_super) {
		__extends(DocumentCreated, _super);
		function DocumentCreated() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(DocumentCreated.prototype, "_className", {
			get: function () {
				return "DocumentCreated";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DocumentCreated.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["saved", "_ReferenceId"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DocumentCreated.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["sections", "body", "contentControls", "properties"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DocumentCreated.prototype, "body", {
			get: function () {
				_throwIfApiNotSupported("DocumentCreated.body", "WordApiHiddenDocument", "1.3", _hostName);
				if (!this._B) {
					this._B=new Word.Body(this.context, _createPropertyObjectPath(this.context, this, "Body", false, false, false));
				}
				return this._B;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DocumentCreated.prototype, "contentControls", {
			get: function () {
				_throwIfApiNotSupported("DocumentCreated.contentControls", "WordApiHiddenDocument", "1.3", _hostName);
				if (!this._C) {
					this._C=new Word.ContentControlCollection(this.context, _createPropertyObjectPath(this.context, this, "ContentControls", true, false, false));
				}
				return this._C;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DocumentCreated.prototype, "properties", {
			get: function () {
				_throwIfApiNotSupported("DocumentCreated.properties", "WordApiHiddenDocument", "1.3", _hostName);
				if (!this._P) {
					this._P=new Word.DocumentProperties(this.context, _createPropertyObjectPath(this.context, this, "Properties", false, false, false));
				}
				return this._P;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DocumentCreated.prototype, "sections", {
			get: function () {
				_throwIfApiNotSupported("DocumentCreated.sections", "WordApiHiddenDocument", "1.3", _hostName);
				if (!this._Se) {
					this._Se=new Word.SectionCollection(this.context, _createPropertyObjectPath(this.context, this, "Sections", true, false, false));
				}
				return this._Se;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DocumentCreated.prototype, "saved", {
			get: function () {
				_throwIfNotLoaded("saved", this._S, _typeDocumentCreated, this._isNull);
				_throwIfApiNotSupported("DocumentCreated.saved", "WordApiHiddenDocument", "1.3", _hostName);
				return this._S;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DocumentCreated.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.m__ReferenceId, _typeDocumentCreated, this._isNull);
				return this.m__ReferenceId;
			},
			enumerable: true,
			configurable: true
		});
		DocumentCreated.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, [], ["body", "properties"], [
				"contentControls",
				"sections"
			]);
		};
		DocumentCreated.prototype.update=function (properties) {
			this._recursivelyUpdate(properties);
		};
		DocumentCreated.prototype.open=function () {
			_createMethodAction(this.context, this, "Open", 1, [], false);
		};
		DocumentCreated.prototype.save=function () {
			_throwIfApiNotSupported("DocumentCreated.save", "WordApiHiddenDocument", "1.3", _hostName);
			_createMethodAction(this.context, this, "Save", 0, [], false);
		};
		DocumentCreated.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1, [], false);
		};
		DocumentCreated.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Saved"])) {
				this._S=obj["Saved"];
			}
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.m__ReferenceId=obj["_ReferenceId"];
			}
			_handleNavigationPropertyResults(this, obj, ["body", "Body", "contentControls", "ContentControls", "properties", "Properties", "sections", "Sections"]);
		};
		DocumentCreated.prototype.load=function (option) {
			return _load(this, option);
		};
		DocumentCreated.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		DocumentCreated.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["_ReferenceId"])) {
				this.m__ReferenceId=value["_ReferenceId"];
			}
		};
		DocumentCreated.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result);
		};
		DocumentCreated.prototype.track=function () {
			this.context.trackedObjects.add(this);
			return this;
		};
		DocumentCreated.prototype.untrack=function () {
			this.context.trackedObjects.remove(this);
			return this;
		};
		DocumentCreated.prototype.toJSON=function () {
			return _toJson(this, {
				"saved": this._S,
			}, {
				"body": this._B,
				"contentControls": this._C,
				"properties": this._P,
				"sections": this._Se,
			});
		};
		DocumentCreated.prototype.ensureUnchanged=function (data) {
			_createEnsureUnchangedAction(this.context, this, data);
			return;
		};
		return DocumentCreated;
	}(OfficeExtension.ClientObject));
	Word.DocumentCreated=DocumentCreated;
	var _typeDocumentProperties="DocumentProperties";
	var DocumentProperties=(function (_super) {
		__extends(DocumentProperties, _super);
		function DocumentProperties() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(DocumentProperties.prototype, "_className", {
			get: function () {
				return "DocumentProperties";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DocumentProperties.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["_ReferenceId", "title", "subject", "author", "keywords", "comments", "template", "lastAuthor", "revisionNumber", "applicationName", "lastPrintDate", "creationDate", "lastSaveTime", "security", "category", "format", "manager", "company"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DocumentProperties.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, true, true, true, true, true, false, false, false, false, false, false, false, false, true, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DocumentProperties.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["customProperties"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DocumentProperties.prototype, "customProperties", {
			get: function () {
				if (!this._Cu) {
					this._Cu=new Word.CustomPropertyCollection(this.context, _createPropertyObjectPath(this.context, this, "CustomProperties", true, false, false));
				}
				return this._Cu;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DocumentProperties.prototype, "applicationName", {
			get: function () {
				_throwIfNotLoaded("applicationName", this._A, _typeDocumentProperties, this._isNull);
				return this._A;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DocumentProperties.prototype, "author", {
			get: function () {
				_throwIfNotLoaded("author", this._Au, _typeDocumentProperties, this._isNull);
				return this._Au;
			},
			set: function (value) {
				this._Au=value;
				_createSetPropertyAction(this.context, this, "Author", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DocumentProperties.prototype, "category", {
			get: function () {
				_throwIfNotLoaded("category", this._C, _typeDocumentProperties, this._isNull);
				return this._C;
			},
			set: function (value) {
				this._C=value;
				_createSetPropertyAction(this.context, this, "Category", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DocumentProperties.prototype, "comments", {
			get: function () {
				_throwIfNotLoaded("comments", this._Co, _typeDocumentProperties, this._isNull);
				return this._Co;
			},
			set: function (value) {
				this._Co=value;
				_createSetPropertyAction(this.context, this, "Comments", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DocumentProperties.prototype, "company", {
			get: function () {
				_throwIfNotLoaded("company", this._Com, _typeDocumentProperties, this._isNull);
				return this._Com;
			},
			set: function (value) {
				this._Com=value;
				_createSetPropertyAction(this.context, this, "Company", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DocumentProperties.prototype, "creationDate", {
			get: function () {
				_throwIfNotLoaded("creationDate", this._Cr, _typeDocumentProperties, this._isNull);
				return this._Cr;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DocumentProperties.prototype, "format", {
			get: function () {
				_throwIfNotLoaded("format", this._F, _typeDocumentProperties, this._isNull);
				return this._F;
			},
			set: function (value) {
				this._F=value;
				_createSetPropertyAction(this.context, this, "Format", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DocumentProperties.prototype, "keywords", {
			get: function () {
				_throwIfNotLoaded("keywords", this._K, _typeDocumentProperties, this._isNull);
				return this._K;
			},
			set: function (value) {
				this._K=value;
				_createSetPropertyAction(this.context, this, "Keywords", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DocumentProperties.prototype, "lastAuthor", {
			get: function () {
				_throwIfNotLoaded("lastAuthor", this._L, _typeDocumentProperties, this._isNull);
				return this._L;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DocumentProperties.prototype, "lastPrintDate", {
			get: function () {
				_throwIfNotLoaded("lastPrintDate", this._La, _typeDocumentProperties, this._isNull);
				return this._La;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DocumentProperties.prototype, "lastSaveTime", {
			get: function () {
				_throwIfNotLoaded("lastSaveTime", this._Las, _typeDocumentProperties, this._isNull);
				return this._Las;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DocumentProperties.prototype, "manager", {
			get: function () {
				_throwIfNotLoaded("manager", this._M, _typeDocumentProperties, this._isNull);
				return this._M;
			},
			set: function (value) {
				this._M=value;
				_createSetPropertyAction(this.context, this, "Manager", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DocumentProperties.prototype, "revisionNumber", {
			get: function () {
				_throwIfNotLoaded("revisionNumber", this._R, _typeDocumentProperties, this._isNull);
				return this._R;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DocumentProperties.prototype, "security", {
			get: function () {
				_throwIfNotLoaded("security", this._S, _typeDocumentProperties, this._isNull);
				return this._S;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DocumentProperties.prototype, "subject", {
			get: function () {
				_throwIfNotLoaded("subject", this._Su, _typeDocumentProperties, this._isNull);
				return this._Su;
			},
			set: function (value) {
				this._Su=value;
				_createSetPropertyAction(this.context, this, "Subject", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DocumentProperties.prototype, "template", {
			get: function () {
				_throwIfNotLoaded("template", this._T, _typeDocumentProperties, this._isNull);
				return this._T;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DocumentProperties.prototype, "title", {
			get: function () {
				_throwIfNotLoaded("title", this._Ti, _typeDocumentProperties, this._isNull);
				return this._Ti;
			},
			set: function (value) {
				this._Ti=value;
				_createSetPropertyAction(this.context, this, "Title", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(DocumentProperties.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.__R, _typeDocumentProperties, this._isNull);
				return this.__R;
			},
			enumerable: true,
			configurable: true
		});
		DocumentProperties.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["title", "subject", "author", "keywords", "comments", "category", "format", "manager", "company"], [], [
				"customProperties"
			]);
		};
		DocumentProperties.prototype.update=function (properties) {
			this._recursivelyUpdate(properties);
		};
		DocumentProperties.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1, [], false);
		};
		DocumentProperties.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["ApplicationName"])) {
				this._A=obj["ApplicationName"];
			}
			if (!_isUndefined(obj["Author"])) {
				this._Au=obj["Author"];
			}
			if (!_isUndefined(obj["Category"])) {
				this._C=obj["Category"];
			}
			if (!_isUndefined(obj["Comments"])) {
				this._Co=obj["Comments"];
			}
			if (!_isUndefined(obj["Company"])) {
				this._Com=obj["Company"];
			}
			if (!_isUndefined(obj["CreationDate"])) {
				this._Cr=_adjustToDateTime(obj["CreationDate"]);
			}
			if (!_isUndefined(obj["Format"])) {
				this._F=obj["Format"];
			}
			if (!_isUndefined(obj["Keywords"])) {
				this._K=obj["Keywords"];
			}
			if (!_isUndefined(obj["LastAuthor"])) {
				this._L=obj["LastAuthor"];
			}
			if (!_isUndefined(obj["LastPrintDate"])) {
				this._La=_adjustToDateTime(obj["LastPrintDate"]);
			}
			if (!_isUndefined(obj["LastSaveTime"])) {
				this._Las=_adjustToDateTime(obj["LastSaveTime"]);
			}
			if (!_isUndefined(obj["Manager"])) {
				this._M=obj["Manager"];
			}
			if (!_isUndefined(obj["RevisionNumber"])) {
				this._R=obj["RevisionNumber"];
			}
			if (!_isUndefined(obj["Security"])) {
				this._S=obj["Security"];
			}
			if (!_isUndefined(obj["Subject"])) {
				this._Su=obj["Subject"];
			}
			if (!_isUndefined(obj["Template"])) {
				this._T=obj["Template"];
			}
			if (!_isUndefined(obj["Title"])) {
				this._Ti=obj["Title"];
			}
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.__R=obj["_ReferenceId"];
			}
			_handleNavigationPropertyResults(this, obj, ["customProperties", "CustomProperties"]);
		};
		DocumentProperties.prototype.load=function (option) {
			return _load(this, option);
		};
		DocumentProperties.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		DocumentProperties.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["_ReferenceId"])) {
				this.__R=value["_ReferenceId"];
			}
		};
		DocumentProperties.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			if (!_isUndefined(obj["CreationDate"])) {
				obj["creationDate"]=_adjustToDateTime(obj["creationDate"]);
			}
			if (!_isUndefined(obj["LastPrintDate"])) {
				obj["lastPrintDate"]=_adjustToDateTime(obj["lastPrintDate"]);
			}
			if (!_isUndefined(obj["LastSaveTime"])) {
				obj["lastSaveTime"]=_adjustToDateTime(obj["lastSaveTime"]);
			}
			_processRetrieveResult(this, value, result);
		};
		DocumentProperties.prototype.track=function () {
			this.context.trackedObjects.add(this);
			return this;
		};
		DocumentProperties.prototype.untrack=function () {
			this.context.trackedObjects.remove(this);
			return this;
		};
		DocumentProperties.prototype.toJSON=function () {
			return _toJson(this, {
				"applicationName": this._A,
				"author": this._Au,
				"category": this._C,
				"comments": this._Co,
				"company": this._Com,
				"creationDate": this._Cr,
				"format": this._F,
				"keywords": this._K,
				"lastAuthor": this._L,
				"lastPrintDate": this._La,
				"lastSaveTime": this._Las,
				"manager": this._M,
				"revisionNumber": this._R,
				"security": this._S,
				"subject": this._Su,
				"template": this._T,
				"title": this._Ti,
			}, {
				"customProperties": this._Cu,
			});
		};
		DocumentProperties.prototype.ensureUnchanged=function (data) {
			_createEnsureUnchangedAction(this.context, this, data);
			return;
		};
		return DocumentProperties;
	}(OfficeExtension.ClientObject));
	Word.DocumentProperties=DocumentProperties;
	var _typeFont="Font";
	var Font=(function (_super) {
		__extends(Font, _super);
		function Font() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Font.prototype, "_className", {
			get: function () {
				return "Font";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Font.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["_ReferenceId", "name", "size", "bold", "italic", "color", "underline", "subscript", "superscript", "strikeThrough", "doubleStrikeThrough", "highlightColor"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Font.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, true, true, true, true, true, true, true, true, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Font.prototype, "bold", {
			get: function () {
				_throwIfNotLoaded("bold", this._B, _typeFont, this._isNull);
				return this._B;
			},
			set: function (value) {
				this._B=value;
				_createSetPropertyAction(this.context, this, "Bold", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Font.prototype, "color", {
			get: function () {
				_throwIfNotLoaded("color", this._C, _typeFont, this._isNull);
				return this._C;
			},
			set: function (value) {
				this._C=value;
				_createSetPropertyAction(this.context, this, "Color", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Font.prototype, "doubleStrikeThrough", {
			get: function () {
				_throwIfNotLoaded("doubleStrikeThrough", this._D, _typeFont, this._isNull);
				return this._D;
			},
			set: function (value) {
				this._D=value;
				_createSetPropertyAction(this.context, this, "DoubleStrikeThrough", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Font.prototype, "highlightColor", {
			get: function () {
				_throwIfNotLoaded("highlightColor", this._H, _typeFont, this._isNull);
				return this._H;
			},
			set: function (value) {
				this._H=value;
				_createSetPropertyAction(this.context, this, "HighlightColor", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Font.prototype, "italic", {
			get: function () {
				_throwIfNotLoaded("italic", this._I, _typeFont, this._isNull);
				return this._I;
			},
			set: function (value) {
				this._I=value;
				_createSetPropertyAction(this.context, this, "Italic", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Font.prototype, "name", {
			get: function () {
				_throwIfNotLoaded("name", this._N, _typeFont, this._isNull);
				return this._N;
			},
			set: function (value) {
				this._N=value;
				_createSetPropertyAction(this.context, this, "Name", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Font.prototype, "size", {
			get: function () {
				_throwIfNotLoaded("size", this._S, _typeFont, this._isNull);
				return this._S;
			},
			set: function (value) {
				this._S=value;
				_createSetPropertyAction(this.context, this, "Size", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Font.prototype, "strikeThrough", {
			get: function () {
				_throwIfNotLoaded("strikeThrough", this._St, _typeFont, this._isNull);
				return this._St;
			},
			set: function (value) {
				this._St=value;
				_createSetPropertyAction(this.context, this, "StrikeThrough", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Font.prototype, "subscript", {
			get: function () {
				_throwIfNotLoaded("subscript", this._Su, _typeFont, this._isNull);
				return this._Su;
			},
			set: function (value) {
				this._Su=value;
				_createSetPropertyAction(this.context, this, "Subscript", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Font.prototype, "superscript", {
			get: function () {
				_throwIfNotLoaded("superscript", this._Sup, _typeFont, this._isNull);
				return this._Sup;
			},
			set: function (value) {
				this._Sup=value;
				_createSetPropertyAction(this.context, this, "Superscript", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Font.prototype, "underline", {
			get: function () {
				_throwIfNotLoaded("underline", this._U, _typeFont, this._isNull);
				return this._U;
			},
			set: function (value) {
				this._U=value;
				_createSetPropertyAction(this.context, this, "Underline", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Font.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.__R, _typeFont, this._isNull);
				return this.__R;
			},
			enumerable: true,
			configurable: true
		});
		Font.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["name", "size", "bold", "italic", "color", "underline", "subscript", "superscript", "strikeThrough", "doubleStrikeThrough", "highlightColor"], [], []);
		};
		Font.prototype.update=function (properties) {
			this._recursivelyUpdate(properties);
		};
		Font.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1, [], false);
		};
		Font.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Bold"])) {
				this._B=obj["Bold"];
			}
			if (!_isUndefined(obj["Color"])) {
				this._C=obj["Color"];
			}
			if (!_isUndefined(obj["DoubleStrikeThrough"])) {
				this._D=obj["DoubleStrikeThrough"];
			}
			if (!_isUndefined(obj["HighlightColor"])) {
				this._H=obj["HighlightColor"];
			}
			if (!_isUndefined(obj["Italic"])) {
				this._I=obj["Italic"];
			}
			if (!_isUndefined(obj["Name"])) {
				this._N=obj["Name"];
			}
			if (!_isUndefined(obj["Size"])) {
				this._S=obj["Size"];
			}
			if (!_isUndefined(obj["StrikeThrough"])) {
				this._St=obj["StrikeThrough"];
			}
			if (!_isUndefined(obj["Subscript"])) {
				this._Su=obj["Subscript"];
			}
			if (!_isUndefined(obj["Superscript"])) {
				this._Sup=obj["Superscript"];
			}
			if (!_isUndefined(obj["Underline"])) {
				this._U=obj["Underline"];
			}
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.__R=obj["_ReferenceId"];
			}
		};
		Font.prototype.load=function (option) {
			return _load(this, option);
		};
		Font.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		Font.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["_ReferenceId"])) {
				this.__R=value["_ReferenceId"];
			}
		};
		Font.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result);
		};
		Font.prototype.track=function () {
			this.context.trackedObjects.add(this);
			return this;
		};
		Font.prototype.untrack=function () {
			this.context.trackedObjects.remove(this);
			return this;
		};
		Font.prototype.toJSON=function () {
			return _toJson(this, {
				"bold": this._B,
				"color": this._C,
				"doubleStrikeThrough": this._D,
				"highlightColor": this._H,
				"italic": this._I,
				"name": this._N,
				"size": this._S,
				"strikeThrough": this._St,
				"subscript": this._Su,
				"superscript": this._Sup,
				"underline": this._U,
			}, {});
		};
		Font.prototype.ensureUnchanged=function (data) {
			_createEnsureUnchangedAction(this.context, this, data);
			return;
		};
		return Font;
	}(OfficeExtension.ClientObject));
	Word.Font=Font;
	var _typeInlinePicture="InlinePicture";
	var InlinePicture=(function (_super) {
		__extends(InlinePicture, _super);
		function InlinePicture() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(InlinePicture.prototype, "_className", {
			get: function () {
				return "InlinePicture";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePicture.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["_Id", "_ReferenceId", "altTextDescription", "altTextTitle", "height", "hyperlink", "lockAspectRatio", "width"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePicture.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, false, true, true, true, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePicture.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["parentContentControl", "paragraph", "parentTableCell", "parentTable", "parentContentControlOrNullObject", "parentTableCellOrNullObject", "parentTableOrNullObject"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePicture.prototype, "paragraph", {
			get: function () {
				_throwIfApiNotSupported("InlinePicture.paragraph", _defaultApiSetName, "1.2", _hostName);
				if (!this._P) {
					this._P=new Word.Paragraph(this.context, _createPropertyObjectPath(this.context, this, "Paragraph", false, false, false));
				}
				return this._P;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePicture.prototype, "parentContentControl", {
			get: function () {
				if (!this._Pa) {
					this._Pa=new Word.ContentControl(this.context, _createPropertyObjectPath(this.context, this, "ParentContentControl", false, false, false));
				}
				return this._Pa;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePicture.prototype, "parentContentControlOrNullObject", {
			get: function () {
				_throwIfApiNotSupported("InlinePicture.parentContentControlOrNullObject", _defaultApiSetName, "1.3", _hostName);
				if (!this._Par) {
					this._Par=new Word.ContentControl(this.context, _createPropertyObjectPath(this.context, this, "ParentContentControlOrNullObject", false, false, false));
				}
				return this._Par;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePicture.prototype, "parentTable", {
			get: function () {
				_throwIfApiNotSupported("InlinePicture.parentTable", _defaultApiSetName, "1.3", _hostName);
				if (!this._Pare) {
					this._Pare=new Word.Table(this.context, _createPropertyObjectPath(this.context, this, "ParentTable", false, false, false));
				}
				return this._Pare;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePicture.prototype, "parentTableCell", {
			get: function () {
				_throwIfApiNotSupported("InlinePicture.parentTableCell", _defaultApiSetName, "1.3", _hostName);
				if (!this._Paren) {
					this._Paren=new Word.TableCell(this.context, _createPropertyObjectPath(this.context, this, "ParentTableCell", false, false, false));
				}
				return this._Paren;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePicture.prototype, "parentTableCellOrNullObject", {
			get: function () {
				_throwIfApiNotSupported("InlinePicture.parentTableCellOrNullObject", _defaultApiSetName, "1.3", _hostName);
				if (!this._Parent) {
					this._Parent=new Word.TableCell(this.context, _createPropertyObjectPath(this.context, this, "ParentTableCellOrNullObject", false, false, false));
				}
				return this._Parent;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePicture.prototype, "parentTableOrNullObject", {
			get: function () {
				_throwIfApiNotSupported("InlinePicture.parentTableOrNullObject", _defaultApiSetName, "1.3", _hostName);
				if (!this._ParentT) {
					this._ParentT=new Word.Table(this.context, _createPropertyObjectPath(this.context, this, "ParentTableOrNullObject", false, false, false));
				}
				return this._ParentT;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePicture.prototype, "altTextDescription", {
			get: function () {
				_throwIfNotLoaded("altTextDescription", this._A, _typeInlinePicture, this._isNull);
				return this._A;
			},
			set: function (value) {
				this._A=value;
				_createSetPropertyAction(this.context, this, "AltTextDescription", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePicture.prototype, "altTextTitle", {
			get: function () {
				_throwIfNotLoaded("altTextTitle", this._Al, _typeInlinePicture, this._isNull);
				return this._Al;
			},
			set: function (value) {
				this._Al=value;
				_createSetPropertyAction(this.context, this, "AltTextTitle", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePicture.prototype, "height", {
			get: function () {
				_throwIfNotLoaded("height", this._H, _typeInlinePicture, this._isNull);
				return this._H;
			},
			set: function (value) {
				this._H=value;
				_createSetPropertyAction(this.context, this, "Height", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePicture.prototype, "hyperlink", {
			get: function () {
				_throwIfNotLoaded("hyperlink", this._Hy, _typeInlinePicture, this._isNull);
				return this._Hy;
			},
			set: function (value) {
				this._Hy=value;
				_createSetPropertyAction(this.context, this, "Hyperlink", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePicture.prototype, "lockAspectRatio", {
			get: function () {
				_throwIfNotLoaded("lockAspectRatio", this._L, _typeInlinePicture, this._isNull);
				return this._L;
			},
			set: function (value) {
				this._L=value;
				_createSetPropertyAction(this.context, this, "LockAspectRatio", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePicture.prototype, "width", {
			get: function () {
				_throwIfNotLoaded("width", this._W, _typeInlinePicture, this._isNull);
				return this._W;
			},
			set: function (value) {
				this._W=value;
				_createSetPropertyAction(this.context, this, "Width", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePicture.prototype, "_Id", {
			get: function () {
				_throwIfNotLoaded("_Id", this.__I, _typeInlinePicture, this._isNull);
				return this.__I;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePicture.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.__R, _typeInlinePicture, this._isNull);
				return this.__R;
			},
			enumerable: true,
			configurable: true
		});
		InlinePicture.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["altTextDescription", "altTextTitle", "height", "hyperlink", "lockAspectRatio", "width"], [], [
				"paragraph",
				"parentContentControl",
				"parentContentControlOrNullObject",
				"parentTable",
				"parentTableCell",
				"parentTableCellOrNullObject",
				"parentTableOrNullObject"
			]);
		};
		InlinePicture.prototype.update=function (properties) {
			this._recursivelyUpdate(properties);
		};
		InlinePicture.prototype.delete=function () {
			_throwIfApiNotSupported("InlinePicture.delete", _defaultApiSetName, "1.2", _hostName);
			_createMethodAction(this.context, this, "Delete", 0, [], false);
		};
		InlinePicture.prototype.getBase64ImageSrc=function () {
			var action=_createMethodAction(this.context, this, "GetBase64ImageSrc", 1, [], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		InlinePicture.prototype.getNext=function () {
			_throwIfApiNotSupported("InlinePicture.getNext", _defaultApiSetName, "1.3", _hostName);
			return new Word.InlinePicture(this.context, _createMethodObjectPath(this.context, this, "GetNext", 1, [], false, false, null, false));
		};
		InlinePicture.prototype.getNextOrNullObject=function () {
			_throwIfApiNotSupported("InlinePicture.getNextOrNullObject", _defaultApiSetName, "1.3", _hostName);
			return new Word.InlinePicture(this.context, _createMethodObjectPath(this.context, this, "GetNextOrNullObject", 1, [], false, false, null, false));
		};
		InlinePicture.prototype.getRange=function (rangeLocation) {
			_throwIfApiNotSupported("InlinePicture.getRange", _defaultApiSetName, "1.3", _hostName);
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "GetRange", 1, [rangeLocation], false, false, null, false));
		};
		InlinePicture.prototype.insertBreak=function (breakType, insertLocation) {
			_throwIfApiNotSupported("InlinePicture.insertBreak", _defaultApiSetName, "1.2", _hostName);
			_createMethodAction(this.context, this, "InsertBreak", 0, [breakType, insertLocation], false);
		};
		InlinePicture.prototype.insertContentControl=function () {
			return new Word.ContentControl(this.context, _createMethodObjectPath(this.context, this, "InsertContentControl", 0, [], false, true, null, false));
		};
		InlinePicture.prototype.insertFileFromBase64=function (base64File, insertLocation) {
			_throwIfApiNotSupported("InlinePicture.insertFileFromBase64", _defaultApiSetName, "1.2", _hostName);
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertFileFromBase64", 0, [base64File, insertLocation], false, true, null, false));
		};
		InlinePicture.prototype.insertHtml=function (html, insertLocation) {
			_throwIfApiNotSupported("InlinePicture.insertHtml", _defaultApiSetName, "1.2", _hostName);
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertHtml", 0, [html, insertLocation], false, true, null, false));
		};
		InlinePicture.prototype.insertInlinePictureFromBase64=function (base64EncodedImage, insertLocation) {
			_throwIfApiNotSupported("InlinePicture.insertInlinePictureFromBase64", _defaultApiSetName, "1.2", _hostName);
			return new Word.InlinePicture(this.context, _createMethodObjectPath(this.context, this, "InsertInlinePictureFromBase64", 0, [base64EncodedImage, insertLocation], false, true, null, false));
		};
		InlinePicture.prototype.insertOoxml=function (ooxml, insertLocation) {
			_throwIfApiNotSupported("InlinePicture.insertOoxml", _defaultApiSetName, "1.2", _hostName);
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertOoxml", 0, [ooxml, insertLocation], false, true, null, false));
		};
		InlinePicture.prototype.insertParagraph=function (paragraphText, insertLocation) {
			_throwIfApiNotSupported("InlinePicture.insertParagraph", _defaultApiSetName, "1.2", _hostName);
			return new Word.Paragraph(this.context, _createMethodObjectPath(this.context, this, "InsertParagraph", 0, [paragraphText, insertLocation], false, true, null, false));
		};
		InlinePicture.prototype.insertText=function (text, insertLocation) {
			_throwIfApiNotSupported("InlinePicture.insertText", _defaultApiSetName, "1.2", _hostName);
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertText", 0, [text, insertLocation], false, true, null, false));
		};
		InlinePicture.prototype.select=function (selectionMode) {
			_throwIfApiNotSupported("InlinePicture.select", _defaultApiSetName, "1.2", _hostName);
			_createMethodAction(this.context, this, "Select", 1, [selectionMode], false);
		};
		InlinePicture.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1, [], false);
		};
		InlinePicture.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["AltTextDescription"])) {
				this._A=obj["AltTextDescription"];
			}
			if (!_isUndefined(obj["AltTextTitle"])) {
				this._Al=obj["AltTextTitle"];
			}
			if (!_isUndefined(obj["Height"])) {
				this._H=obj["Height"];
			}
			if (!_isUndefined(obj["Hyperlink"])) {
				this._Hy=obj["Hyperlink"];
			}
			if (!_isUndefined(obj["LockAspectRatio"])) {
				this._L=obj["LockAspectRatio"];
			}
			if (!_isUndefined(obj["Width"])) {
				this._W=obj["Width"];
			}
			if (!_isUndefined(obj["_Id"])) {
				this.__I=obj["_Id"];
			}
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.__R=obj["_ReferenceId"];
			}
			_handleNavigationPropertyResults(this, obj, ["paragraph", "Paragraph", "parentContentControl", "ParentContentControl", "parentContentControlOrNullObject", "ParentContentControlOrNullObject", "parentTable", "ParentTable", "parentTableCell", "ParentTableCell", "parentTableCellOrNullObject", "ParentTableCellOrNullObject", "parentTableOrNullObject", "ParentTableOrNullObject"]);
		};
		InlinePicture.prototype.load=function (option) {
			return _load(this, option);
		};
		InlinePicture.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		InlinePicture.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["_ReferenceId"])) {
				this.__R=value["_ReferenceId"];
			}
			if (!_isUndefined(value["_Id"])) {
				this.__I=value["_Id"];
			}
		};
		InlinePicture.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result);
		};
		InlinePicture.prototype.track=function () {
			this.context.trackedObjects.add(this);
			return this;
		};
		InlinePicture.prototype.untrack=function () {
			this.context.trackedObjects.remove(this);
			return this;
		};
		InlinePicture.prototype.toJSON=function () {
			return _toJson(this, {
				"altTextDescription": this._A,
				"altTextTitle": this._Al,
				"height": this._H,
				"hyperlink": this._Hy,
				"lockAspectRatio": this._L,
				"width": this._W,
			}, {});
		};
		InlinePicture.prototype.ensureUnchanged=function (data) {
			_createEnsureUnchangedAction(this.context, this, data);
			return;
		};
		return InlinePicture;
	}(OfficeExtension.ClientObject));
	Word.InlinePicture=InlinePicture;
	var _typeInlinePictureCollection="InlinePictureCollection";
	var InlinePictureCollection=(function (_super) {
		__extends(InlinePictureCollection, _super);
		function InlinePictureCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(InlinePictureCollection.prototype, "_className", {
			get: function () {
				return "InlinePictureCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePictureCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePictureCollection.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["_ReferenceId"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePictureCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, _typeInlinePictureCollection, this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(InlinePictureCollection.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.__R, _typeInlinePictureCollection, this._isNull);
				return this.__R;
			},
			enumerable: true,
			configurable: true
		});
		InlinePictureCollection.prototype.getFirst=function () {
			_throwIfApiNotSupported("InlinePictureCollection.getFirst", _defaultApiSetName, "1.3", _hostName);
			return new Word.InlinePicture(this.context, _createMethodObjectPath(this.context, this, "GetFirst", 1, [], false, false, null, false));
		};
		InlinePictureCollection.prototype.getFirstOrNullObject=function () {
			_throwIfApiNotSupported("InlinePictureCollection.getFirstOrNullObject", _defaultApiSetName, "1.3", _hostName);
			return new Word.InlinePicture(this.context, _createMethodObjectPath(this.context, this, "GetFirstOrNullObject", 1, [], false, false, null, false));
		};
		InlinePictureCollection.prototype._GetItem=function (index) {
			return new Word.InlinePicture(this.context, _createIndexerObjectPath(this.context, this, [index]));
		};
		InlinePictureCollection.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1, [], false);
		};
		InlinePictureCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.__R=obj["_ReferenceId"];
			}
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Word.InlinePicture(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		InlinePictureCollection.prototype.load=function (option) {
			return _load(this, option);
		};
		InlinePictureCollection.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		InlinePictureCollection.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["_ReferenceId"])) {
				this.__R=value["_ReferenceId"];
			}
		};
		InlinePictureCollection.prototype._handleRetrieveResult=function (value, result) {
			var _this=this;
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result, function (childItemData, index) { return new Word.InlinePicture(_this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, _this.context, _this, childItemData, index)); });
		};
		InlinePictureCollection.prototype.track=function () {
			this.context.trackedObjects.add(this);
			return this;
		};
		InlinePictureCollection.prototype.untrack=function () {
			this.context.trackedObjects.remove(this);
			return this;
		};
		InlinePictureCollection.prototype.toJSON=function () {
			return _toJson(this, {}, {}, this.m__items);
		};
		return InlinePictureCollection;
	}(OfficeExtension.ClientObject));
	Word.InlinePictureCollection=InlinePictureCollection;
	var _typeList="List";
	var List=(function (_super) {
		__extends(List, _super);
		function List() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(List.prototype, "_className", {
			get: function () {
				return "List";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(List.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["id", "_ReferenceId", "levelTypes", "levelExistences"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(List.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["paragraphs"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(List.prototype, "paragraphs", {
			get: function () {
				if (!this._P) {
					this._P=new Word.ParagraphCollection(this.context, _createPropertyObjectPath(this.context, this, "Paragraphs", true, false, false));
				}
				return this._P;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(List.prototype, "id", {
			get: function () {
				_throwIfNotLoaded("id", this._I, _typeList, this._isNull);
				return this._I;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(List.prototype, "levelExistences", {
			get: function () {
				_throwIfNotLoaded("levelExistences", this._L, _typeList, this._isNull);
				return this._L;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(List.prototype, "levelTypes", {
			get: function () {
				_throwIfNotLoaded("levelTypes", this._Le, _typeList, this._isNull);
				return this._Le;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(List.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.__R, _typeList, this._isNull);
				return this.__R;
			},
			enumerable: true,
			configurable: true
		});
		List.prototype.getLevelParagraphs=function (level) {
			return new Word.ParagraphCollection(this.context, _createMethodObjectPath(this.context, this, "GetLevelParagraphs", 1, [level], true, false, null, false));
		};
		List.prototype.getLevelString=function (level) {
			var action=_createMethodAction(this.context, this, "GetLevelString", 1, [level], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		List.prototype.insertParagraph=function (paragraphText, insertLocation) {
			return new Word.Paragraph(this.context, _createMethodObjectPath(this.context, this, "InsertParagraph", 0, [paragraphText, insertLocation], false, true, null, false));
		};
		List.prototype.setLevelAlignment=function (level, alignment) {
			_createMethodAction(this.context, this, "SetLevelAlignment", 0, [level, alignment], false);
		};
		List.prototype.setLevelBullet=function (level, listBullet, charCode, fontName) {
			_createMethodAction(this.context, this, "SetLevelBullet", 0, [level, listBullet, charCode, fontName], false);
		};
		List.prototype.setLevelIndents=function (level, textIndent, bulletNumberPictureIndent) {
			_createMethodAction(this.context, this, "SetLevelIndents", 0, [level, textIndent, bulletNumberPictureIndent], false);
		};
		List.prototype.setLevelNumbering=function (level, listNumbering, formatString) {
			_createMethodAction(this.context, this, "SetLevelNumbering", 0, [level, listNumbering, formatString], false);
		};
		List.prototype.setLevelStartingNumber=function (level, startingNumber) {
			_createMethodAction(this.context, this, "SetLevelStartingNumber", 0, [level, startingNumber], false);
		};
		List.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1, [], false);
		};
		List.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Id"])) {
				this._I=obj["Id"];
			}
			if (!_isUndefined(obj["LevelExistences"])) {
				this._L=obj["LevelExistences"];
			}
			if (!_isUndefined(obj["LevelTypes"])) {
				this._Le=obj["LevelTypes"];
			}
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.__R=obj["_ReferenceId"];
			}
			_handleNavigationPropertyResults(this, obj, ["paragraphs", "Paragraphs"]);
		};
		List.prototype.load=function (option) {
			return _load(this, option);
		};
		List.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		List.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["_ReferenceId"])) {
				this.__R=value["_ReferenceId"];
			}
			if (!_isUndefined(value["Id"])) {
				this._I=value["Id"];
			}
		};
		List.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result);
		};
		List.prototype.track=function () {
			this.context.trackedObjects.add(this);
			return this;
		};
		List.prototype.untrack=function () {
			this.context.trackedObjects.remove(this);
			return this;
		};
		List.prototype.toJSON=function () {
			return _toJson(this, {
				"id": this._I,
				"levelExistences": this._L,
				"levelTypes": this._Le,
			}, {
				"paragraphs": this._P,
			});
		};
		List.prototype.ensureUnchanged=function (data) {
			_createEnsureUnchangedAction(this.context, this, data);
			return;
		};
		return List;
	}(OfficeExtension.ClientObject));
	Word.List=List;
	var _typeListCollection="ListCollection";
	var ListCollection=(function (_super) {
		__extends(ListCollection, _super);
		function ListCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ListCollection.prototype, "_className", {
			get: function () {
				return "ListCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ListCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ListCollection.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["_ReferenceId"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ListCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, _typeListCollection, this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ListCollection.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.__R, _typeListCollection, this._isNull);
				return this.__R;
			},
			enumerable: true,
			configurable: true
		});
		ListCollection.prototype.getById=function (id) {
			return new Word.List(this.context, _createMethodObjectPath(this.context, this, "GetById", 1, [id], false, false, null, false));
		};
		ListCollection.prototype.getByIdOrNullObject=function (id) {
			return new Word.List(this.context, _createMethodObjectPath(this.context, this, "GetByIdOrNullObject", 1, [id], false, false, null, false));
		};
		ListCollection.prototype.getFirst=function () {
			return new Word.List(this.context, _createMethodObjectPath(this.context, this, "GetFirst", 1, [], false, false, null, false));
		};
		ListCollection.prototype.getFirstOrNullObject=function () {
			return new Word.List(this.context, _createMethodObjectPath(this.context, this, "GetFirstOrNullObject", 1, [], false, false, null, false));
		};
		ListCollection.prototype.getItem=function (index) {
			return new Word.List(this.context, _createIndexerObjectPath(this.context, this, [index]));
		};
		ListCollection.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1, [], false);
		};
		ListCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.__R=obj["_ReferenceId"];
			}
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Word.List(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		ListCollection.prototype.load=function (option) {
			return _load(this, option);
		};
		ListCollection.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		ListCollection.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["_ReferenceId"])) {
				this.__R=value["_ReferenceId"];
			}
		};
		ListCollection.prototype._handleRetrieveResult=function (value, result) {
			var _this=this;
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result, function (childItemData, index) { return new Word.List(_this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, _this.context, _this, childItemData, index)); });
		};
		ListCollection.prototype.track=function () {
			this.context.trackedObjects.add(this);
			return this;
		};
		ListCollection.prototype.untrack=function () {
			this.context.trackedObjects.remove(this);
			return this;
		};
		ListCollection.prototype.toJSON=function () {
			return _toJson(this, {}, {}, this.m__items);
		};
		return ListCollection;
	}(OfficeExtension.ClientObject));
	Word.ListCollection=ListCollection;
	var _typeListItem="ListItem";
	var ListItem=(function (_super) {
		__extends(ListItem, _super);
		function ListItem() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ListItem.prototype, "_className", {
			get: function () {
				return "ListItem";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ListItem.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["_ReferenceId", "siblingIndex", "listString", "level"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ListItem.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, false, false, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ListItem.prototype, "level", {
			get: function () {
				_throwIfNotLoaded("level", this._L, _typeListItem, this._isNull);
				return this._L;
			},
			set: function (value) {
				this._L=value;
				_createSetPropertyAction(this.context, this, "Level", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ListItem.prototype, "listString", {
			get: function () {
				_throwIfNotLoaded("listString", this._Li, _typeListItem, this._isNull);
				return this._Li;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ListItem.prototype, "siblingIndex", {
			get: function () {
				_throwIfNotLoaded("siblingIndex", this._S, _typeListItem, this._isNull);
				return this._S;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ListItem.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.__R, _typeListItem, this._isNull);
				return this.__R;
			},
			enumerable: true,
			configurable: true
		});
		ListItem.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["level"], [], []);
		};
		ListItem.prototype.update=function (properties) {
			this._recursivelyUpdate(properties);
		};
		ListItem.prototype.getAncestor=function (parentOnly) {
			return new Word.Paragraph(this.context, _createMethodObjectPath(this.context, this, "GetAncestor", 1, [parentOnly], false, false, null, false));
		};
		ListItem.prototype.getAncestorOrNullObject=function (parentOnly) {
			return new Word.Paragraph(this.context, _createMethodObjectPath(this.context, this, "GetAncestorOrNullObject", 1, [parentOnly], false, false, null, false));
		};
		ListItem.prototype.getDescendants=function (directChildrenOnly) {
			return new Word.ParagraphCollection(this.context, _createMethodObjectPath(this.context, this, "GetDescendants", 1, [directChildrenOnly], true, false, null, false));
		};
		ListItem.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1, [], false);
		};
		ListItem.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Level"])) {
				this._L=obj["Level"];
			}
			if (!_isUndefined(obj["ListString"])) {
				this._Li=obj["ListString"];
			}
			if (!_isUndefined(obj["SiblingIndex"])) {
				this._S=obj["SiblingIndex"];
			}
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.__R=obj["_ReferenceId"];
			}
		};
		ListItem.prototype.load=function (option) {
			return _load(this, option);
		};
		ListItem.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		ListItem.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["_ReferenceId"])) {
				this.__R=value["_ReferenceId"];
			}
		};
		ListItem.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result);
		};
		ListItem.prototype.track=function () {
			this.context.trackedObjects.add(this);
			return this;
		};
		ListItem.prototype.untrack=function () {
			this.context.trackedObjects.remove(this);
			return this;
		};
		ListItem.prototype.toJSON=function () {
			return _toJson(this, {
				"level": this._L,
				"listString": this._Li,
				"siblingIndex": this._S,
			}, {});
		};
		ListItem.prototype.ensureUnchanged=function (data) {
			_createEnsureUnchangedAction(this.context, this, data);
			return;
		};
		return ListItem;
	}(OfficeExtension.ClientObject));
	Word.ListItem=ListItem;
	var _typeParagraph="Paragraph";
	var Paragraph=(function (_super) {
		__extends(Paragraph, _super);
		function Paragraph() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Paragraph.prototype, "_className", {
			get: function () {
				return "Paragraph";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["_Id", "_ReferenceId", "style", "alignment", "firstLineIndent", "leftIndent", "rightIndent", "lineSpacing", "outlineLevel", "spaceBefore", "spaceAfter", "lineUnitBefore", "lineUnitAfter", "text", "isListItem", "tableNestingLevel", "isLastParagraph", "styleBuiltIn"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, false, true, true, true, true, true, true, true, true, true, true, true, false, false, false, false, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["font", "contentControls", "parentContentControl", "inlinePictures", "parentBody", "list", "parentTableCell", "parentTable", "listItem", "parentContentControlOrNullObject", "parentTableCellOrNullObject", "parentTableOrNullObject", "listOrNullObject", "listItemOrNullObject"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "contentControls", {
			get: function () {
				if (!this._C) {
					this._C=new Word.ContentControlCollection(this.context, _createPropertyObjectPath(this.context, this, "ContentControls", true, false, false));
				}
				return this._C;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "font", {
			get: function () {
				if (!this._Fo) {
					this._Fo=new Word.Font(this.context, _createPropertyObjectPath(this.context, this, "Font", false, false, false));
				}
				return this._Fo;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "inlinePictures", {
			get: function () {
				if (!this._I) {
					this._I=new Word.InlinePictureCollection(this.context, _createPropertyObjectPath(this.context, this, "InlinePictures", true, false, false));
				}
				return this._I;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "list", {
			get: function () {
				_throwIfApiNotSupported("Paragraph.list", _defaultApiSetName, "1.3", _hostName);
				if (!this._Lis) {
					this._Lis=new Word.List(this.context, _createPropertyObjectPath(this.context, this, "List", false, false, false));
				}
				return this._Lis;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "listItem", {
			get: function () {
				_throwIfApiNotSupported("Paragraph.listItem", _defaultApiSetName, "1.3", _hostName);
				if (!this._List) {
					this._List=new Word.ListItem(this.context, _createPropertyObjectPath(this.context, this, "ListItem", false, false, false));
				}
				return this._List;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "listItemOrNullObject", {
			get: function () {
				_throwIfApiNotSupported("Paragraph.listItemOrNullObject", _defaultApiSetName, "1.3", _hostName);
				if (!this._ListI) {
					this._ListI=new Word.ListItem(this.context, _createPropertyObjectPath(this.context, this, "ListItemOrNullObject", false, false, false));
				}
				return this._ListI;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "listOrNullObject", {
			get: function () {
				_throwIfApiNotSupported("Paragraph.listOrNullObject", _defaultApiSetName, "1.3", _hostName);
				if (!this._ListO) {
					this._ListO=new Word.List(this.context, _createPropertyObjectPath(this.context, this, "ListOrNullObject", false, false, false));
				}
				return this._ListO;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "parentBody", {
			get: function () {
				_throwIfApiNotSupported("Paragraph.parentBody", _defaultApiSetName, "1.3", _hostName);
				if (!this._P) {
					this._P=new Word.Body(this.context, _createPropertyObjectPath(this.context, this, "ParentBody", false, false, false));
				}
				return this._P;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "parentContentControl", {
			get: function () {
				if (!this._Pa) {
					this._Pa=new Word.ContentControl(this.context, _createPropertyObjectPath(this.context, this, "ParentContentControl", false, false, false));
				}
				return this._Pa;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "parentContentControlOrNullObject", {
			get: function () {
				_throwIfApiNotSupported("Paragraph.parentContentControlOrNullObject", _defaultApiSetName, "1.3", _hostName);
				if (!this._Par) {
					this._Par=new Word.ContentControl(this.context, _createPropertyObjectPath(this.context, this, "ParentContentControlOrNullObject", false, false, false));
				}
				return this._Par;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "parentTable", {
			get: function () {
				_throwIfApiNotSupported("Paragraph.parentTable", _defaultApiSetName, "1.3", _hostName);
				if (!this._Pare) {
					this._Pare=new Word.Table(this.context, _createPropertyObjectPath(this.context, this, "ParentTable", false, false, false));
				}
				return this._Pare;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "parentTableCell", {
			get: function () {
				_throwIfApiNotSupported("Paragraph.parentTableCell", _defaultApiSetName, "1.3", _hostName);
				if (!this._Paren) {
					this._Paren=new Word.TableCell(this.context, _createPropertyObjectPath(this.context, this, "ParentTableCell", false, false, false));
				}
				return this._Paren;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "parentTableCellOrNullObject", {
			get: function () {
				_throwIfApiNotSupported("Paragraph.parentTableCellOrNullObject", _defaultApiSetName, "1.3", _hostName);
				if (!this._Parent) {
					this._Parent=new Word.TableCell(this.context, _createPropertyObjectPath(this.context, this, "ParentTableCellOrNullObject", false, false, false));
				}
				return this._Parent;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "parentTableOrNullObject", {
			get: function () {
				_throwIfApiNotSupported("Paragraph.parentTableOrNullObject", _defaultApiSetName, "1.3", _hostName);
				if (!this._ParentT) {
					this._ParentT=new Word.Table(this.context, _createPropertyObjectPath(this.context, this, "ParentTableOrNullObject", false, false, false));
				}
				return this._ParentT;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "alignment", {
			get: function () {
				_throwIfNotLoaded("alignment", this._A, _typeParagraph, this._isNull);
				return this._A;
			},
			set: function (value) {
				this._A=value;
				_createSetPropertyAction(this.context, this, "Alignment", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "firstLineIndent", {
			get: function () {
				_throwIfNotLoaded("firstLineIndent", this._F, _typeParagraph, this._isNull);
				return this._F;
			},
			set: function (value) {
				this._F=value;
				_createSetPropertyAction(this.context, this, "FirstLineIndent", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "isLastParagraph", {
			get: function () {
				_throwIfNotLoaded("isLastParagraph", this._Is, _typeParagraph, this._isNull);
				_throwIfApiNotSupported("Paragraph.isLastParagraph", _defaultApiSetName, "1.3", _hostName);
				return this._Is;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "isListItem", {
			get: function () {
				_throwIfNotLoaded("isListItem", this._IsL, _typeParagraph, this._isNull);
				_throwIfApiNotSupported("Paragraph.isListItem", _defaultApiSetName, "1.3", _hostName);
				return this._IsL;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "leftIndent", {
			get: function () {
				_throwIfNotLoaded("leftIndent", this._L, _typeParagraph, this._isNull);
				return this._L;
			},
			set: function (value) {
				this._L=value;
				_createSetPropertyAction(this.context, this, "LeftIndent", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "lineSpacing", {
			get: function () {
				_throwIfNotLoaded("lineSpacing", this._Li, _typeParagraph, this._isNull);
				return this._Li;
			},
			set: function (value) {
				this._Li=value;
				_createSetPropertyAction(this.context, this, "LineSpacing", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "lineUnitAfter", {
			get: function () {
				_throwIfNotLoaded("lineUnitAfter", this._Lin, _typeParagraph, this._isNull);
				return this._Lin;
			},
			set: function (value) {
				this._Lin=value;
				_createSetPropertyAction(this.context, this, "LineUnitAfter", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "lineUnitBefore", {
			get: function () {
				_throwIfNotLoaded("lineUnitBefore", this._Line, _typeParagraph, this._isNull);
				return this._Line;
			},
			set: function (value) {
				this._Line=value;
				_createSetPropertyAction(this.context, this, "LineUnitBefore", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "outlineLevel", {
			get: function () {
				_throwIfNotLoaded("outlineLevel", this._O, _typeParagraph, this._isNull);
				return this._O;
			},
			set: function (value) {
				this._O=value;
				_createSetPropertyAction(this.context, this, "OutlineLevel", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "rightIndent", {
			get: function () {
				_throwIfNotLoaded("rightIndent", this._R, _typeParagraph, this._isNull);
				return this._R;
			},
			set: function (value) {
				this._R=value;
				_createSetPropertyAction(this.context, this, "RightIndent", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "spaceAfter", {
			get: function () {
				_throwIfNotLoaded("spaceAfter", this._S, _typeParagraph, this._isNull);
				return this._S;
			},
			set: function (value) {
				this._S=value;
				_createSetPropertyAction(this.context, this, "SpaceAfter", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "spaceBefore", {
			get: function () {
				_throwIfNotLoaded("spaceBefore", this._Sp, _typeParagraph, this._isNull);
				return this._Sp;
			},
			set: function (value) {
				this._Sp=value;
				_createSetPropertyAction(this.context, this, "SpaceBefore", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "style", {
			get: function () {
				_throwIfNotLoaded("style", this._St, _typeParagraph, this._isNull);
				return this._St;
			},
			set: function (value) {
				this._St=value;
				_createSetPropertyAction(this.context, this, "Style", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "styleBuiltIn", {
			get: function () {
				_throwIfNotLoaded("styleBuiltIn", this._Sty, _typeParagraph, this._isNull);
				_throwIfApiNotSupported("Paragraph.styleBuiltIn", _defaultApiSetName, "1.3", _hostName);
				return this._Sty;
			},
			set: function (value) {
				this._Sty=value;
				_createSetPropertyAction(this.context, this, "StyleBuiltIn", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "tableNestingLevel", {
			get: function () {
				_throwIfNotLoaded("tableNestingLevel", this._T, _typeParagraph, this._isNull);
				_throwIfApiNotSupported("Paragraph.tableNestingLevel", _defaultApiSetName, "1.3", _hostName);
				return this._T;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "text", {
			get: function () {
				_throwIfNotLoaded("text", this._Te, _typeParagraph, this._isNull);
				return this._Te;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "_Id", {
			get: function () {
				_throwIfNotLoaded("_Id", this.__I, _typeParagraph, this._isNull);
				return this.__I;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Paragraph.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.__R, _typeParagraph, this._isNull);
				return this.__R;
			},
			enumerable: true,
			configurable: true
		});
		Paragraph.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["style", "alignment", "firstLineIndent", "leftIndent", "rightIndent", "lineSpacing", "outlineLevel", "spaceBefore", "spaceAfter", "lineUnitBefore", "lineUnitAfter", "styleBuiltIn"], ["font", "listItem", "listItemOrNullObject"], [
				"contentControls",
				"inlinePictures",
				"list",
				"listOrNullObject",
				"parentBody",
				"parentContentControl",
				"parentContentControlOrNullObject",
				"parentTable",
				"parentTableCell",
				"parentTableCellOrNullObject",
				"parentTableOrNullObject"
			]);
		};
		Paragraph.prototype.update=function (properties) {
			this._recursivelyUpdate(properties);
		};
		Paragraph.prototype.attachToList=function (listId, level) {
			_throwIfApiNotSupported("Paragraph.attachToList", _defaultApiSetName, "1.3", _hostName);
			return new Word.List(this.context, _createMethodObjectPath(this.context, this, "AttachToList", 0, [listId, level], false, false, null, false));
		};
		Paragraph.prototype.clear=function () {
			_createMethodAction(this.context, this, "Clear", 0, [], false);
		};
		Paragraph.prototype.delete=function () {
			_createMethodAction(this.context, this, "Delete", 0, [], false);
		};
		Paragraph.prototype.detachFromList=function () {
			_throwIfApiNotSupported("Paragraph.detachFromList", _defaultApiSetName, "1.3", _hostName);
			_createMethodAction(this.context, this, "DetachFromList", 0, [], false);
		};
		Paragraph.prototype.getHtml=function () {
			var action=_createMethodAction(this.context, this, "GetHtml", 1, [], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		Paragraph.prototype.getMetadata=function (propertyGUID) {
			var action=_createMethodAction(this.context, this, "GetMetadata", 1, [propertyGUID], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		Paragraph.prototype.getNext=function () {
			_throwIfApiNotSupported("Paragraph.getNext", _defaultApiSetName, "1.3", _hostName);
			return new Word.Paragraph(this.context, _createMethodObjectPath(this.context, this, "GetNext", 1, [], false, false, null, false));
		};
		Paragraph.prototype.getNextOrNullObject=function () {
			_throwIfApiNotSupported("Paragraph.getNextOrNullObject", _defaultApiSetName, "1.3", _hostName);
			return new Word.Paragraph(this.context, _createMethodObjectPath(this.context, this, "GetNextOrNullObject", 1, [], false, false, null, false));
		};
		Paragraph.prototype.getOoxml=function () {
			var action=_createMethodAction(this.context, this, "GetOoxml", 1, [], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		Paragraph.prototype.getPrevious=function () {
			_throwIfApiNotSupported("Paragraph.getPrevious", _defaultApiSetName, "1.3", _hostName);
			return new Word.Paragraph(this.context, _createMethodObjectPath(this.context, this, "GetPrevious", 1, [], false, false, null, false));
		};
		Paragraph.prototype.getPreviousOrNullObject=function () {
			_throwIfApiNotSupported("Paragraph.getPreviousOrNullObject", _defaultApiSetName, "1.3", _hostName);
			return new Word.Paragraph(this.context, _createMethodObjectPath(this.context, this, "GetPreviousOrNullObject", 1, [], false, false, null, false));
		};
		Paragraph.prototype.getRange=function (rangeLocation) {
			_throwIfApiNotSupported("Paragraph.getRange", _defaultApiSetName, "1.3", _hostName);
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "GetRange", 1, [rangeLocation], false, false, null, false));
		};
		Paragraph.prototype.getTextRanges=function (endingMarks, trimSpacing) {
			_throwIfApiNotSupported("Paragraph.getTextRanges", _defaultApiSetName, "1.3", _hostName);
			return new Word.RangeCollection(this.context, _createMethodObjectPath(this.context, this, "GetTextRanges", 1, [endingMarks, trimSpacing], true, false, null, false));
		};
		Paragraph.prototype.insertBreak=function (breakType, insertLocation) {
			_createMethodAction(this.context, this, "InsertBreak", 0, [breakType, insertLocation], false);
		};
		Paragraph.prototype.insertContentControl=function () {
			return new Word.ContentControl(this.context, _createMethodObjectPath(this.context, this, "InsertContentControl", 0, [], false, true, null, false));
		};
		Paragraph.prototype.insertFileFromBase64=function (base64File, insertLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertFileFromBase64", 0, [base64File, insertLocation], false, true, null, false));
		};
		Paragraph.prototype.insertHtml=function (html, insertLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertHtml", 0, [html, insertLocation], false, true, null, false));
		};
		Paragraph.prototype.insertInlinePictureFromBase64=function (base64EncodedImage, insertLocation) {
			return new Word.InlinePicture(this.context, _createMethodObjectPath(this.context, this, "InsertInlinePictureFromBase64", 0, [base64EncodedImage, insertLocation], false, true, null, false));
		};
		Paragraph.prototype.insertOoxml=function (ooxml, insertLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertOoxml", 0, [ooxml, insertLocation], false, true, null, false));
		};
		Paragraph.prototype.insertParagraph=function (paragraphText, insertLocation) {
			return new Word.Paragraph(this.context, _createMethodObjectPath(this.context, this, "InsertParagraph", 0, [paragraphText, insertLocation], false, true, null, false));
		};
		Paragraph.prototype.insertTable=function (rowCount, columnCount, insertLocation, values) {
			_throwIfApiNotSupported("Paragraph.insertTable", _defaultApiSetName, "1.3", _hostName);
			return new Word.Table(this.context, _createMethodObjectPath(this.context, this, "InsertTable", 0, [rowCount, columnCount, insertLocation, values], false, true, null, false));
		};
		Paragraph.prototype.insertText=function (text, insertLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertText", 0, [text, insertLocation], false, true, null, false));
		};
		Paragraph.prototype.search=function (searchText, searchOptions) {
			searchOptions=_normalizeSearchOptions(this.context, searchOptions);
			return new Word.RangeCollection(this.context, _createMethodObjectPath(this.context, this, "Search", 1, [searchText, searchOptions], true, false, null, false));
		};
		Paragraph.prototype.select=function (selectionMode) {
			_createMethodAction(this.context, this, "Select", 1, [selectionMode], false);
		};
		Paragraph.prototype.setMetadata=function (propertyGUID, propertyValue) {
			_createMethodAction(this.context, this, "SetMetadata", 0, [propertyGUID, propertyValue], false);
		};
		Paragraph.prototype.split=function (delimiters, trimDelimiters, trimSpacing) {
			_throwIfApiNotSupported("Paragraph.split", _defaultApiSetName, "1.3", _hostName);
			return new Word.RangeCollection(this.context, _createMethodObjectPath(this.context, this, "Split", 1, [delimiters, trimDelimiters, trimSpacing], true, false, null, false));
		};
		Paragraph.prototype.startNewList=function () {
			_throwIfApiNotSupported("Paragraph.startNewList", _defaultApiSetName, "1.3", _hostName);
			return new Word.List(this.context, _createMethodObjectPath(this.context, this, "StartNewList", 0, [], false, false, null, false));
		};
		Paragraph.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1, [], false);
		};
		Paragraph.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Alignment"])) {
				this._A=obj["Alignment"];
			}
			if (!_isUndefined(obj["FirstLineIndent"])) {
				this._F=obj["FirstLineIndent"];
			}
			if (!_isUndefined(obj["IsLastParagraph"])) {
				this._Is=obj["IsLastParagraph"];
			}
			if (!_isUndefined(obj["IsListItem"])) {
				this._IsL=obj["IsListItem"];
			}
			if (!_isUndefined(obj["LeftIndent"])) {
				this._L=obj["LeftIndent"];
			}
			if (!_isUndefined(obj["LineSpacing"])) {
				this._Li=obj["LineSpacing"];
			}
			if (!_isUndefined(obj["LineUnitAfter"])) {
				this._Lin=obj["LineUnitAfter"];
			}
			if (!_isUndefined(obj["LineUnitBefore"])) {
				this._Line=obj["LineUnitBefore"];
			}
			if (!_isUndefined(obj["OutlineLevel"])) {
				this._O=obj["OutlineLevel"];
			}
			if (!_isUndefined(obj["RightIndent"])) {
				this._R=obj["RightIndent"];
			}
			if (!_isUndefined(obj["SpaceAfter"])) {
				this._S=obj["SpaceAfter"];
			}
			if (!_isUndefined(obj["SpaceBefore"])) {
				this._Sp=obj["SpaceBefore"];
			}
			if (!_isUndefined(obj["Style"])) {
				this._St=obj["Style"];
			}
			if (!_isUndefined(obj["StyleBuiltIn"])) {
				this._Sty=obj["StyleBuiltIn"];
			}
			if (!_isUndefined(obj["TableNestingLevel"])) {
				this._T=obj["TableNestingLevel"];
			}
			if (!_isUndefined(obj["Text"])) {
				this._Te=obj["Text"];
			}
			if (!_isUndefined(obj["_Id"])) {
				this.__I=obj["_Id"];
			}
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.__R=obj["_ReferenceId"];
			}
			_handleNavigationPropertyResults(this, obj, ["contentControls", "ContentControls", "font", "Font", "inlinePictures", "InlinePictures", "list", "List", "listItem", "ListItem", "listItemOrNullObject", "ListItemOrNullObject", "listOrNullObject", "ListOrNullObject", "parentBody", "ParentBody", "parentContentControl", "ParentContentControl", "parentContentControlOrNullObject", "ParentContentControlOrNullObject", "parentTable", "ParentTable", "parentTableCell", "ParentTableCell", "parentTableCellOrNullObject", "ParentTableCellOrNullObject", "parentTableOrNullObject", "ParentTableOrNullObject"]);
		};
		Paragraph.prototype.load=function (option) {
			return _load(this, option);
		};
		Paragraph.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		Paragraph.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["_ReferenceId"])) {
				this.__R=value["_ReferenceId"];
			}
			if (!_isUndefined(value["_Id"])) {
				this.__I=value["_Id"];
			}
		};
		Paragraph.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result);
		};
		Paragraph.prototype.track=function () {
			this.context.trackedObjects.add(this);
			return this;
		};
		Paragraph.prototype.untrack=function () {
			this.context.trackedObjects.remove(this);
			return this;
		};
		Paragraph.prototype.toJSON=function () {
			return _toJson(this, {
				"alignment": this._A,
				"firstLineIndent": this._F,
				"isLastParagraph": this._Is,
				"isListItem": this._IsL,
				"leftIndent": this._L,
				"lineSpacing": this._Li,
				"lineUnitAfter": this._Lin,
				"lineUnitBefore": this._Line,
				"outlineLevel": this._O,
				"rightIndent": this._R,
				"spaceAfter": this._S,
				"spaceBefore": this._Sp,
				"style": this._St,
				"styleBuiltIn": this._Sty,
				"tableNestingLevel": this._T,
				"text": this._Te,
			}, {
				"font": this._Fo,
				"inlinePictures": this._I,
				"listItem": this._List,
				"listItemOrNullObject": this._ListI,
			});
		};
		Paragraph.prototype.ensureUnchanged=function (data) {
			_createEnsureUnchangedAction(this.context, this, data);
			return;
		};
		return Paragraph;
	}(OfficeExtension.ClientObject));
	Word.Paragraph=Paragraph;
	var _typeParagraphCollection="ParagraphCollection";
	var ParagraphCollection=(function (_super) {
		__extends(ParagraphCollection, _super);
		function ParagraphCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(ParagraphCollection.prototype, "_className", {
			get: function () {
				return "ParagraphCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ParagraphCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ParagraphCollection.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["_ReferenceId"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ParagraphCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, _typeParagraphCollection, this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(ParagraphCollection.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.__R, _typeParagraphCollection, this._isNull);
				return this.__R;
			},
			enumerable: true,
			configurable: true
		});
		ParagraphCollection.prototype.getFirst=function () {
			_throwIfApiNotSupported("ParagraphCollection.getFirst", _defaultApiSetName, "1.3", _hostName);
			return new Word.Paragraph(this.context, _createMethodObjectPath(this.context, this, "GetFirst", 1, [], false, false, null, false));
		};
		ParagraphCollection.prototype.getFirstOrNullObject=function () {
			_throwIfApiNotSupported("ParagraphCollection.getFirstOrNullObject", _defaultApiSetName, "1.3", _hostName);
			return new Word.Paragraph(this.context, _createMethodObjectPath(this.context, this, "GetFirstOrNullObject", 1, [], false, false, null, false));
		};
		ParagraphCollection.prototype.getLast=function () {
			_throwIfApiNotSupported("ParagraphCollection.getLast", _defaultApiSetName, "1.3", _hostName);
			return new Word.Paragraph(this.context, _createMethodObjectPath(this.context, this, "GetLast", 1, [], false, false, null, false));
		};
		ParagraphCollection.prototype.getLastOrNullObject=function () {
			_throwIfApiNotSupported("ParagraphCollection.getLastOrNullObject", _defaultApiSetName, "1.3", _hostName);
			return new Word.Paragraph(this.context, _createMethodObjectPath(this.context, this, "GetLastOrNullObject", 1, [], false, false, null, false));
		};
		ParagraphCollection.prototype._GetItem=function (index) {
			return new Word.Paragraph(this.context, _createIndexerObjectPath(this.context, this, [index]));
		};
		ParagraphCollection.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1, [], false);
		};
		ParagraphCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.__R=obj["_ReferenceId"];
			}
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Word.Paragraph(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		ParagraphCollection.prototype.load=function (option) {
			return _load(this, option);
		};
		ParagraphCollection.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		ParagraphCollection.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["_ReferenceId"])) {
				this.__R=value["_ReferenceId"];
			}
		};
		ParagraphCollection.prototype._handleRetrieveResult=function (value, result) {
			var _this=this;
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result, function (childItemData, index) { return new Word.Paragraph(_this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, _this.context, _this, childItemData, index)); });
		};
		ParagraphCollection.prototype.track=function () {
			this.context.trackedObjects.add(this);
			return this;
		};
		ParagraphCollection.prototype.untrack=function () {
			this.context.trackedObjects.remove(this);
			return this;
		};
		ParagraphCollection.prototype.toJSON=function () {
			return _toJson(this, {}, {}, this.m__items);
		};
		return ParagraphCollection;
	}(OfficeExtension.ClientObject));
	Word.ParagraphCollection=ParagraphCollection;
	var _typeRange="Range";
	var Range=(function (_super) {
		__extends(Range, _super);
		function Range() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Range.prototype, "_className", {
			get: function () {
				return "Range";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["_Id", "_ReferenceId", "style", "text", "isEmpty", "hyperlink", "styleBuiltIn"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, false, true, false, false, true, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["font", "paragraphs", "contentControls", "parentContentControl", "inlinePictures", "lists", "tables", "parentTableCell", "parentTable", "parentBody", "parentContentControlOrNullObject", "parentTableCellOrNullObject", "parentTableOrNullObject"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "contentControls", {
			get: function () {
				if (!this._C) {
					this._C=new Word.ContentControlCollection(this.context, _createPropertyObjectPath(this.context, this, "ContentControls", true, false, false));
				}
				return this._C;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "font", {
			get: function () {
				if (!this._F) {
					this._F=new Word.Font(this.context, _createPropertyObjectPath(this.context, this, "Font", false, false, false));
				}
				return this._F;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "inlinePictures", {
			get: function () {
				_throwIfApiNotSupported("Range.inlinePictures", _defaultApiSetName, "1.2", _hostName);
				if (!this._I) {
					this._I=new Word.InlinePictureCollection(this.context, _createPropertyObjectPath(this.context, this, "InlinePictures", true, false, false));
				}
				return this._I;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "lists", {
			get: function () {
				_throwIfApiNotSupported("Range.lists", _defaultApiSetName, "1.3", _hostName);
				if (!this._L) {
					this._L=new Word.ListCollection(this.context, _createPropertyObjectPath(this.context, this, "Lists", true, false, false));
				}
				return this._L;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "paragraphs", {
			get: function () {
				if (!this._P) {
					this._P=new Word.ParagraphCollection(this.context, _createPropertyObjectPath(this.context, this, "Paragraphs", true, false, false));
				}
				return this._P;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "parentBody", {
			get: function () {
				_throwIfApiNotSupported("Range.parentBody", _defaultApiSetName, "1.3", _hostName);
				if (!this._Pa) {
					this._Pa=new Word.Body(this.context, _createPropertyObjectPath(this.context, this, "ParentBody", false, false, false));
				}
				return this._Pa;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "parentContentControl", {
			get: function () {
				if (!this._Par) {
					this._Par=new Word.ContentControl(this.context, _createPropertyObjectPath(this.context, this, "ParentContentControl", false, false, false));
				}
				return this._Par;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "parentContentControlOrNullObject", {
			get: function () {
				_throwIfApiNotSupported("Range.parentContentControlOrNullObject", _defaultApiSetName, "1.3", _hostName);
				if (!this._Pare) {
					this._Pare=new Word.ContentControl(this.context, _createPropertyObjectPath(this.context, this, "ParentContentControlOrNullObject", false, false, false));
				}
				return this._Pare;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "parentTable", {
			get: function () {
				_throwIfApiNotSupported("Range.parentTable", _defaultApiSetName, "1.3", _hostName);
				if (!this._Paren) {
					this._Paren=new Word.Table(this.context, _createPropertyObjectPath(this.context, this, "ParentTable", false, false, false));
				}
				return this._Paren;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "parentTableCell", {
			get: function () {
				_throwIfApiNotSupported("Range.parentTableCell", _defaultApiSetName, "1.3", _hostName);
				if (!this._Parent) {
					this._Parent=new Word.TableCell(this.context, _createPropertyObjectPath(this.context, this, "ParentTableCell", false, false, false));
				}
				return this._Parent;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "parentTableCellOrNullObject", {
			get: function () {
				_throwIfApiNotSupported("Range.parentTableCellOrNullObject", _defaultApiSetName, "1.3", _hostName);
				if (!this._ParentT) {
					this._ParentT=new Word.TableCell(this.context, _createPropertyObjectPath(this.context, this, "ParentTableCellOrNullObject", false, false, false));
				}
				return this._ParentT;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "parentTableOrNullObject", {
			get: function () {
				_throwIfApiNotSupported("Range.parentTableOrNullObject", _defaultApiSetName, "1.3", _hostName);
				if (!this._ParentTa) {
					this._ParentTa=new Word.Table(this.context, _createPropertyObjectPath(this.context, this, "ParentTableOrNullObject", false, false, false));
				}
				return this._ParentTa;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "tables", {
			get: function () {
				_throwIfApiNotSupported("Range.tables", _defaultApiSetName, "1.3", _hostName);
				if (!this._T) {
					this._T=new Word.TableCollection(this.context, _createPropertyObjectPath(this.context, this, "Tables", true, false, false));
				}
				return this._T;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "hyperlink", {
			get: function () {
				_throwIfNotLoaded("hyperlink", this._H, _typeRange, this._isNull);
				_throwIfApiNotSupported("Range.hyperlink", _defaultApiSetName, "1.3", _hostName);
				return this._H;
			},
			set: function (value) {
				this._H=value;
				_createSetPropertyAction(this.context, this, "Hyperlink", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "isEmpty", {
			get: function () {
				_throwIfNotLoaded("isEmpty", this._Is, _typeRange, this._isNull);
				_throwIfApiNotSupported("Range.isEmpty", _defaultApiSetName, "1.3", _hostName);
				return this._Is;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "style", {
			get: function () {
				_throwIfNotLoaded("style", this._S, _typeRange, this._isNull);
				return this._S;
			},
			set: function (value) {
				this._S=value;
				_createSetPropertyAction(this.context, this, "Style", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "styleBuiltIn", {
			get: function () {
				_throwIfNotLoaded("styleBuiltIn", this._St, _typeRange, this._isNull);
				_throwIfApiNotSupported("Range.styleBuiltIn", _defaultApiSetName, "1.3", _hostName);
				return this._St;
			},
			set: function (value) {
				this._St=value;
				_createSetPropertyAction(this.context, this, "StyleBuiltIn", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "text", {
			get: function () {
				_throwIfNotLoaded("text", this._Te, _typeRange, this._isNull);
				return this._Te;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "_Id", {
			get: function () {
				_throwIfNotLoaded("_Id", this.__I, _typeRange, this._isNull);
				return this.__I;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Range.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.__R, _typeRange, this._isNull);
				return this.__R;
			},
			enumerable: true,
			configurable: true
		});
		Range.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["style", "hyperlink", "styleBuiltIn"], ["font"], [
				"contentControls",
				"inlinePictures",
				"lists",
				"paragraphs",
				"parentBody",
				"parentContentControl",
				"parentContentControlOrNullObject",
				"parentTable",
				"parentTableCell",
				"parentTableCellOrNullObject",
				"parentTableOrNullObject",
				"tables"
			]);
		};
		Range.prototype.update=function (properties) {
			this._recursivelyUpdate(properties);
		};
		Range.prototype.clear=function () {
			_createMethodAction(this.context, this, "Clear", 0, [], false);
		};
		Range.prototype.compareLocationWith=function (range) {
			_throwIfApiNotSupported("Range.compareLocationWith", _defaultApiSetName, "1.3", _hostName);
			var action=_createMethodAction(this.context, this, "CompareLocationWith", 1, [range], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		Range.prototype.delete=function () {
			_createMethodAction(this.context, this, "Delete", 0, [], false);
		};
		Range.prototype.expandTo=function (range) {
			_throwIfApiNotSupported("Range.expandTo", _defaultApiSetName, "1.3", _hostName);
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "ExpandTo", 0, [range], false, false, null, false));
		};
		Range.prototype.expandToOrNullObject=function (range) {
			_throwIfApiNotSupported("Range.expandToOrNullObject", _defaultApiSetName, "1.3", _hostName);
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "ExpandToOrNullObject", 0, [range], false, false, null, false));
		};
		Range.prototype.getHtml=function () {
			var action=_createMethodAction(this.context, this, "GetHtml", 1, [], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		Range.prototype.getHyperlinkRanges=function () {
			_throwIfApiNotSupported("Range.getHyperlinkRanges", _defaultApiSetName, "1.3", _hostName);
			return new Word.RangeCollection(this.context, _createMethodObjectPath(this.context, this, "GetHyperlinkRanges", 1, [], true, false, null, false));
		};
		Range.prototype.getNextTextRange=function (endingMarks, trimSpacing) {
			_throwIfApiNotSupported("Range.getNextTextRange", _defaultApiSetName, "1.3", _hostName);
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "GetNextTextRange", 1, [endingMarks, trimSpacing], false, false, null, false));
		};
		Range.prototype.getNextTextRangeOrNullObject=function (endingMarks, trimSpacing) {
			_throwIfApiNotSupported("Range.getNextTextRangeOrNullObject", _defaultApiSetName, "1.3", _hostName);
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "GetNextTextRangeOrNullObject", 1, [endingMarks, trimSpacing], false, false, null, false));
		};
		Range.prototype.getOoxml=function () {
			var action=_createMethodAction(this.context, this, "GetOoxml", 1, [], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		Range.prototype.getRange=function (rangeLocation) {
			_throwIfApiNotSupported("Range.getRange", _defaultApiSetName, "1.3", _hostName);
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "GetRange", 1, [rangeLocation], false, false, null, false));
		};
		Range.prototype.getTextRanges=function (endingMarks, trimSpacing) {
			_throwIfApiNotSupported("Range.getTextRanges", _defaultApiSetName, "1.3", _hostName);
			return new Word.RangeCollection(this.context, _createMethodObjectPath(this.context, this, "GetTextRanges", 1, [endingMarks, trimSpacing], true, false, null, false));
		};
		Range.prototype.insertBreak=function (breakType, insertLocation) {
			_createMethodAction(this.context, this, "InsertBreak", 0, [breakType, insertLocation], false);
		};
		Range.prototype.insertContentControl=function () {
			return new Word.ContentControl(this.context, _createMethodObjectPath(this.context, this, "InsertContentControl", 0, [], false, true, null, false));
		};
		Range.prototype.insertFileFromBase64=function (base64File, insertLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertFileFromBase64", 0, [base64File, insertLocation], false, true, null, false));
		};
		Range.prototype.insertHtml=function (html, insertLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertHtml", 0, [html, insertLocation], false, true, null, false));
		};
		Range.prototype.insertInlinePictureFromBase64=function (base64EncodedImage, insertLocation) {
			_throwIfApiNotSupported("Range.insertInlinePictureFromBase64", _defaultApiSetName, "1.2", _hostName);
			return new Word.InlinePicture(this.context, _createMethodObjectPath(this.context, this, "InsertInlinePictureFromBase64", 0, [base64EncodedImage, insertLocation], false, true, null, false));
		};
		Range.prototype.insertOoxml=function (ooxml, insertLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertOoxml", 0, [ooxml, insertLocation], false, true, null, false));
		};
		Range.prototype.insertParagraph=function (paragraphText, insertLocation) {
			return new Word.Paragraph(this.context, _createMethodObjectPath(this.context, this, "InsertParagraph", 0, [paragraphText, insertLocation], false, true, null, false));
		};
		Range.prototype.insertTable=function (rowCount, columnCount, insertLocation, values) {
			_throwIfApiNotSupported("Range.insertTable", _defaultApiSetName, "1.3", _hostName);
			return new Word.Table(this.context, _createMethodObjectPath(this.context, this, "InsertTable", 0, [rowCount, columnCount, insertLocation, values], false, true, null, false));
		};
		Range.prototype.insertTapObjectFromFileContent=function (fileContent, objectType, contentFormat) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertTapObjectFromFileContent", 0, [fileContent, objectType, contentFormat], false, true, null, true));
		};
		Range.prototype.insertText=function (text, insertLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "InsertText", 0, [text, insertLocation], false, true, null, false));
		};
		Range.prototype.intersectWith=function (range) {
			_throwIfApiNotSupported("Range.intersectWith", _defaultApiSetName, "1.3", _hostName);
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "IntersectWith", 0, [range], false, false, null, false));
		};
		Range.prototype.intersectWithOrNullObject=function (range) {
			_throwIfApiNotSupported("Range.intersectWithOrNullObject", _defaultApiSetName, "1.3", _hostName);
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "IntersectWithOrNullObject", 0, [range], false, false, null, false));
		};
		Range.prototype.search=function (searchText, searchOptions) {
			searchOptions=_normalizeSearchOptions(this.context, searchOptions);
			return new Word.RangeCollection(this.context, _createMethodObjectPath(this.context, this, "Search", 1, [searchText, searchOptions], true, false, null, false));
		};
		Range.prototype.select=function (selectionMode) {
			_createMethodAction(this.context, this, "Select", 1, [selectionMode], false);
		};
		Range.prototype.split=function (delimiters, multiParagraphs, trimDelimiters, trimSpacing) {
			_throwIfApiNotSupported("Range.split", _defaultApiSetName, "1.3", _hostName);
			return new Word.RangeCollection(this.context, _createMethodObjectPath(this.context, this, "Split", 1, [delimiters, multiParagraphs, trimDelimiters, trimSpacing], true, false, null, false));
		};
		Range.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1, [], false);
		};
		Range.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Hyperlink"])) {
				this._H=obj["Hyperlink"];
			}
			if (!_isUndefined(obj["IsEmpty"])) {
				this._Is=obj["IsEmpty"];
			}
			if (!_isUndefined(obj["Style"])) {
				this._S=obj["Style"];
			}
			if (!_isUndefined(obj["StyleBuiltIn"])) {
				this._St=obj["StyleBuiltIn"];
			}
			if (!_isUndefined(obj["Text"])) {
				this._Te=obj["Text"];
			}
			if (!_isUndefined(obj["_Id"])) {
				this.__I=obj["_Id"];
			}
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.__R=obj["_ReferenceId"];
			}
			_handleNavigationPropertyResults(this, obj, ["contentControls", "ContentControls", "font", "Font", "inlinePictures", "InlinePictures", "lists", "Lists", "paragraphs", "Paragraphs", "parentBody", "ParentBody", "parentContentControl", "ParentContentControl", "parentContentControlOrNullObject", "ParentContentControlOrNullObject", "parentTable", "ParentTable", "parentTableCell", "ParentTableCell", "parentTableCellOrNullObject", "ParentTableCellOrNullObject", "parentTableOrNullObject", "ParentTableOrNullObject", "tables", "Tables"]);
		};
		Range.prototype.load=function (option) {
			return _load(this, option);
		};
		Range.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		Range.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["_ReferenceId"])) {
				this.__R=value["_ReferenceId"];
			}
			if (!_isUndefined(value["_Id"])) {
				this.__I=value["_Id"];
			}
		};
		Range.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result);
		};
		Range.prototype.track=function () {
			this.context.trackedObjects.add(this);
			return this;
		};
		Range.prototype.untrack=function () {
			this.context.trackedObjects.remove(this);
			return this;
		};
		Range.prototype.toJSON=function () {
			return _toJson(this, {
				"hyperlink": this._H,
				"isEmpty": this._Is,
				"style": this._S,
				"styleBuiltIn": this._St,
				"text": this._Te,
			}, {
				"font": this._F,
				"inlinePictures": this._I,
			});
		};
		Range.prototype.ensureUnchanged=function (data) {
			_createEnsureUnchangedAction(this.context, this, data);
			return;
		};
		return Range;
	}(OfficeExtension.ClientObject));
	Word.Range=Range;
	var _typeRangeCollection="RangeCollection";
	var RangeCollection=(function (_super) {
		__extends(RangeCollection, _super);
		function RangeCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(RangeCollection.prototype, "_className", {
			get: function () {
				return "RangeCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeCollection.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["_ReferenceId"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, _typeRangeCollection, this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RangeCollection.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.__R, _typeRangeCollection, this._isNull);
				return this.__R;
			},
			enumerable: true,
			configurable: true
		});
		RangeCollection.prototype.getFirst=function () {
			_throwIfApiNotSupported("RangeCollection.getFirst", _defaultApiSetName, "1.3", _hostName);
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "GetFirst", 1, [], false, false, null, false));
		};
		RangeCollection.prototype.getFirstOrNullObject=function () {
			_throwIfApiNotSupported("RangeCollection.getFirstOrNullObject", _defaultApiSetName, "1.3", _hostName);
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "GetFirstOrNullObject", 1, [], false, false, null, false));
		};
		RangeCollection.prototype._GetItem=function (index) {
			return new Word.Range(this.context, _createIndexerObjectPath(this.context, this, [index]));
		};
		RangeCollection.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1, [], false);
		};
		RangeCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.__R=obj["_ReferenceId"];
			}
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Word.Range(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		RangeCollection.prototype.load=function (option) {
			return _load(this, option);
		};
		RangeCollection.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		RangeCollection.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["_ReferenceId"])) {
				this.__R=value["_ReferenceId"];
			}
		};
		RangeCollection.prototype._handleRetrieveResult=function (value, result) {
			var _this=this;
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result, function (childItemData, index) { return new Word.Range(_this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, _this.context, _this, childItemData, index)); });
		};
		RangeCollection.prototype.track=function () {
			this.context.trackedObjects.add(this);
			return this;
		};
		RangeCollection.prototype.untrack=function () {
			this.context.trackedObjects.remove(this);
			return this;
		};
		RangeCollection.prototype.toJSON=function () {
			return _toJson(this, {}, {}, this.m__items);
		};
		return RangeCollection;
	}(OfficeExtension.ClientObject));
	Word.RangeCollection=RangeCollection;
	var _typeSearchOptions="SearchOptions";
	var SearchOptions=(function (_super) {
		__extends(SearchOptions, _super);
		function SearchOptions() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(SearchOptions.prototype, "_className", {
			get: function () {
				return "SearchOptions";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(SearchOptions.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["ignorePunct", "ignoreSpace", "matchCase", "matchPrefix", "matchSuffix", "matchWildcards", "matchWholeWord"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(SearchOptions.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [true, true, true, true, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(SearchOptions.prototype, "matchWildCards", {
			get: function () {
				_throwIfNotLoaded("matchWildCards", this.m_matchWildcards);
				return this.m_matchWildcards;
			},
			set: function (value) {
				this.m_matchWildcards=value;
				_createSetPropertyAction(this.context, this, "MatchWildCards", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(SearchOptions.prototype, "ignorePunct", {
			get: function () {
				_throwIfNotLoaded("ignorePunct", this._I, _typeSearchOptions, this._isNull);
				return this._I;
			},
			set: function (value) {
				this._I=value;
				_createSetPropertyAction(this.context, this, "IgnorePunct", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(SearchOptions.prototype, "ignoreSpace", {
			get: function () {
				_throwIfNotLoaded("ignoreSpace", this._Ig, _typeSearchOptions, this._isNull);
				return this._Ig;
			},
			set: function (value) {
				this._Ig=value;
				_createSetPropertyAction(this.context, this, "IgnoreSpace", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(SearchOptions.prototype, "matchCase", {
			get: function () {
				_throwIfNotLoaded("matchCase", this._M, _typeSearchOptions, this._isNull);
				return this._M;
			},
			set: function (value) {
				this._M=value;
				_createSetPropertyAction(this.context, this, "MatchCase", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(SearchOptions.prototype, "matchPrefix", {
			get: function () {
				_throwIfNotLoaded("matchPrefix", this._Ma, _typeSearchOptions, this._isNull);
				return this._Ma;
			},
			set: function (value) {
				this._Ma=value;
				_createSetPropertyAction(this.context, this, "MatchPrefix", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(SearchOptions.prototype, "matchSuffix", {
			get: function () {
				_throwIfNotLoaded("matchSuffix", this._Mat, _typeSearchOptions, this._isNull);
				return this._Mat;
			},
			set: function (value) {
				this._Mat=value;
				_createSetPropertyAction(this.context, this, "MatchSuffix", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(SearchOptions.prototype, "matchWholeWord", {
			get: function () {
				_throwIfNotLoaded("matchWholeWord", this._Matc, _typeSearchOptions, this._isNull);
				return this._Matc;
			},
			set: function (value) {
				this._Matc=value;
				_createSetPropertyAction(this.context, this, "MatchWholeWord", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(SearchOptions.prototype, "matchWildcards", {
			get: function () {
				_throwIfNotLoaded("matchWildcards", this.m_matchWildcards, _typeSearchOptions, this._isNull);
				return this.m_matchWildcards;
			},
			set: function (value) {
				this.m_matchWildcards=value;
				_createSetPropertyAction(this.context, this, "MatchWildcards", value);
			},
			enumerable: true,
			configurable: true
		});
		SearchOptions.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["ignorePunct", "ignoreSpace", "matchCase", "matchPrefix", "matchSuffix", "matchWildcards", "matchWholeWord"], [], []);
		};
		SearchOptions.prototype.update=function (properties) {
			this._recursivelyUpdate(properties);
		};
		SearchOptions.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["IgnorePunct"])) {
				this._I=obj["IgnorePunct"];
			}
			if (!_isUndefined(obj["IgnoreSpace"])) {
				this._Ig=obj["IgnoreSpace"];
			}
			if (!_isUndefined(obj["MatchCase"])) {
				this._M=obj["MatchCase"];
			}
			if (!_isUndefined(obj["MatchPrefix"])) {
				this._Ma=obj["MatchPrefix"];
			}
			if (!_isUndefined(obj["MatchSuffix"])) {
				this._Mat=obj["MatchSuffix"];
			}
			if (!_isUndefined(obj["MatchWholeWord"])) {
				this._Matc=obj["MatchWholeWord"];
			}
			if (!_isUndefined(obj["MatchWildcards"])) {
				this.m_matchWildcards=obj["MatchWildcards"];
			}
		};
		SearchOptions.prototype.load=function (option) {
			return _load(this, option);
		};
		SearchOptions.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		SearchOptions.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result);
		};
		SearchOptions.newObject=function (context) {
			var ret=new Word.SearchOptions(context, _createNewObjectObjectPath(context, "Microsoft.WordServices.SearchOptions", false, false));
			return ret;
		};
		SearchOptions.prototype.toJSON=function () {
			return _toJson(this, {
				"ignorePunct": this._I,
				"ignoreSpace": this._Ig,
				"matchCase": this._M,
				"matchPrefix": this._Ma,
				"matchSuffix": this._Mat,
				"matchWholeWord": this._Matc,
				"matchWildcards": this.m_matchWildcards,
			}, {});
		};
		SearchOptions.prototype.ensureUnchanged=function (data) {
			_createEnsureUnchangedAction(this.context, this, data);
			return;
		};
		return SearchOptions;
	}(OfficeExtension.ClientObject));
	Word.SearchOptions=SearchOptions;
	var _typeSection="Section";
	var Section=(function (_super) {
		__extends(Section, _super);
		function Section() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Section.prototype, "_className", {
			get: function () {
				return "Section";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Section.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["_Id", "_ReferenceId"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Section.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["body"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Section.prototype, "body", {
			get: function () {
				if (!this._B) {
					this._B=new Word.Body(this.context, _createPropertyObjectPath(this.context, this, "Body", false, false, false));
				}
				return this._B;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Section.prototype, "_Id", {
			get: function () {
				_throwIfNotLoaded("_Id", this.__I, _typeSection, this._isNull);
				return this.__I;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Section.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.__R, _typeSection, this._isNull);
				return this.__R;
			},
			enumerable: true,
			configurable: true
		});
		Section.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, [], ["body"], []);
		};
		Section.prototype.update=function (properties) {
			this._recursivelyUpdate(properties);
		};
		Section.prototype.getFooter=function (type) {
			return new Word.Body(this.context, _createMethodObjectPath(this.context, this, "GetFooter", 1, [type], false, true, null, false));
		};
		Section.prototype.getHeader=function (type) {
			return new Word.Body(this.context, _createMethodObjectPath(this.context, this, "GetHeader", 1, [type], false, true, null, false));
		};
		Section.prototype.getNext=function () {
			_throwIfApiNotSupported("Section.getNext", _defaultApiSetName, "1.3", _hostName);
			return new Word.Section(this.context, _createMethodObjectPath(this.context, this, "GetNext", 1, [], false, false, null, false));
		};
		Section.prototype.getNextOrNullObject=function () {
			_throwIfApiNotSupported("Section.getNextOrNullObject", _defaultApiSetName, "1.3", _hostName);
			return new Word.Section(this.context, _createMethodObjectPath(this.context, this, "GetNextOrNullObject", 1, [], false, false, null, false));
		};
		Section.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1, [], false);
		};
		Section.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["_Id"])) {
				this.__I=obj["_Id"];
			}
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.__R=obj["_ReferenceId"];
			}
			_handleNavigationPropertyResults(this, obj, ["body", "Body"]);
		};
		Section.prototype.load=function (option) {
			return _load(this, option);
		};
		Section.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		Section.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["_ReferenceId"])) {
				this.__R=value["_ReferenceId"];
			}
			if (!_isUndefined(value["_Id"])) {
				this.__I=value["_Id"];
			}
		};
		Section.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result);
		};
		Section.prototype.track=function () {
			this.context.trackedObjects.add(this);
			return this;
		};
		Section.prototype.untrack=function () {
			this.context.trackedObjects.remove(this);
			return this;
		};
		Section.prototype.toJSON=function () {
			return _toJson(this, {}, {
				"body": this._B,
			});
		};
		Section.prototype.ensureUnchanged=function (data) {
			_createEnsureUnchangedAction(this.context, this, data);
			return;
		};
		return Section;
	}(OfficeExtension.ClientObject));
	Word.Section=Section;
	var _typeSectionCollection="SectionCollection";
	var SectionCollection=(function (_super) {
		__extends(SectionCollection, _super);
		function SectionCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(SectionCollection.prototype, "_className", {
			get: function () {
				return "SectionCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(SectionCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(SectionCollection.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["_ReferenceId"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(SectionCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, _typeSectionCollection, this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(SectionCollection.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.__R, _typeSectionCollection, this._isNull);
				return this.__R;
			},
			enumerable: true,
			configurable: true
		});
		SectionCollection.prototype.getFirst=function () {
			_throwIfApiNotSupported("SectionCollection.getFirst", _defaultApiSetName, "1.3", _hostName);
			return new Word.Section(this.context, _createMethodObjectPath(this.context, this, "GetFirst", 1, [], false, false, null, false));
		};
		SectionCollection.prototype.getFirstOrNullObject=function () {
			_throwIfApiNotSupported("SectionCollection.getFirstOrNullObject", _defaultApiSetName, "1.3", _hostName);
			return new Word.Section(this.context, _createMethodObjectPath(this.context, this, "GetFirstOrNullObject", 1, [], false, false, null, false));
		};
		SectionCollection.prototype._GetItem=function (index) {
			return new Word.Section(this.context, _createIndexerObjectPath(this.context, this, [index]));
		};
		SectionCollection.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1, [], false);
		};
		SectionCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.__R=obj["_ReferenceId"];
			}
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Word.Section(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		SectionCollection.prototype.load=function (option) {
			return _load(this, option);
		};
		SectionCollection.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		SectionCollection.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["_ReferenceId"])) {
				this.__R=value["_ReferenceId"];
			}
		};
		SectionCollection.prototype._handleRetrieveResult=function (value, result) {
			var _this=this;
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result, function (childItemData, index) { return new Word.Section(_this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, _this.context, _this, childItemData, index)); });
		};
		SectionCollection.prototype.track=function () {
			this.context.trackedObjects.add(this);
			return this;
		};
		SectionCollection.prototype.untrack=function () {
			this.context.trackedObjects.remove(this);
			return this;
		};
		SectionCollection.prototype.toJSON=function () {
			return _toJson(this, {}, {}, this.m__items);
		};
		return SectionCollection;
	}(OfficeExtension.ClientObject));
	Word.SectionCollection=SectionCollection;
	var _typeTable="Table";
	var Table=(function (_super) {
		__extends(Table, _super);
		function Table() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(Table.prototype, "_className", {
			get: function () {
				return "Table";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["_Id", "_ReferenceId", "isUniform", "nestingLevel", "values", "style", "rowCount", "headerRowCount", "styleTotalRow", "styleFirstColumn", "styleLastColumn", "styleBandedRows", "styleBandedColumns", "shadingColor", "horizontalAlignment", "verticalAlignment", "width", "styleBuiltIn", "alignment"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, false, false, false, true, true, false, true, true, true, true, true, true, true, true, true, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["rows", "tables", "parentTableCell", "parentTable", "font", "parentContentControl", "parentBody", "parentTableCellOrNullObject", "parentTableOrNullObject", "parentContentControlOrNullObject"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "font", {
			get: function () {
				if (!this._F) {
					this._F=new Word.Font(this.context, _createPropertyObjectPath(this.context, this, "Font", false, false, false));
				}
				return this._F;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "parentBody", {
			get: function () {
				if (!this._P) {
					this._P=new Word.Body(this.context, _createPropertyObjectPath(this.context, this, "ParentBody", false, false, false));
				}
				return this._P;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "parentContentControl", {
			get: function () {
				if (!this._Pa) {
					this._Pa=new Word.ContentControl(this.context, _createPropertyObjectPath(this.context, this, "ParentContentControl", false, false, false));
				}
				return this._Pa;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "parentContentControlOrNullObject", {
			get: function () {
				if (!this._Par) {
					this._Par=new Word.ContentControl(this.context, _createPropertyObjectPath(this.context, this, "ParentContentControlOrNullObject", false, false, false));
				}
				return this._Par;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "parentTable", {
			get: function () {
				if (!this._Pare) {
					this._Pare=new Word.Table(this.context, _createPropertyObjectPath(this.context, this, "ParentTable", false, false, false));
				}
				return this._Pare;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "parentTableCell", {
			get: function () {
				if (!this._Paren) {
					this._Paren=new Word.TableCell(this.context, _createPropertyObjectPath(this.context, this, "ParentTableCell", false, false, false));
				}
				return this._Paren;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "parentTableCellOrNullObject", {
			get: function () {
				if (!this._Parent) {
					this._Parent=new Word.TableCell(this.context, _createPropertyObjectPath(this.context, this, "ParentTableCellOrNullObject", false, false, false));
				}
				return this._Parent;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "parentTableOrNullObject", {
			get: function () {
				if (!this._ParentT) {
					this._ParentT=new Word.Table(this.context, _createPropertyObjectPath(this.context, this, "ParentTableOrNullObject", false, false, false));
				}
				return this._ParentT;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "rows", {
			get: function () {
				if (!this._Ro) {
					this._Ro=new Word.TableRowCollection(this.context, _createPropertyObjectPath(this.context, this, "Rows", true, false, false));
				}
				return this._Ro;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "tables", {
			get: function () {
				if (!this._T) {
					this._T=new Word.TableCollection(this.context, _createPropertyObjectPath(this.context, this, "Tables", true, false, false));
				}
				return this._T;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "alignment", {
			get: function () {
				_throwIfNotLoaded("alignment", this._A, _typeTable, this._isNull);
				return this._A;
			},
			set: function (value) {
				this._A=value;
				_createSetPropertyAction(this.context, this, "Alignment", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "headerRowCount", {
			get: function () {
				_throwIfNotLoaded("headerRowCount", this._H, _typeTable, this._isNull);
				return this._H;
			},
			set: function (value) {
				this._H=value;
				_createSetPropertyAction(this.context, this, "HeaderRowCount", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "horizontalAlignment", {
			get: function () {
				_throwIfNotLoaded("horizontalAlignment", this._Ho, _typeTable, this._isNull);
				return this._Ho;
			},
			set: function (value) {
				this._Ho=value;
				_createSetPropertyAction(this.context, this, "HorizontalAlignment", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "isUniform", {
			get: function () {
				_throwIfNotLoaded("isUniform", this._I, _typeTable, this._isNull);
				return this._I;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "nestingLevel", {
			get: function () {
				_throwIfNotLoaded("nestingLevel", this._N, _typeTable, this._isNull);
				return this._N;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "rowCount", {
			get: function () {
				_throwIfNotLoaded("rowCount", this._R, _typeTable, this._isNull);
				return this._R;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "shadingColor", {
			get: function () {
				_throwIfNotLoaded("shadingColor", this._S, _typeTable, this._isNull);
				return this._S;
			},
			set: function (value) {
				this._S=value;
				_createSetPropertyAction(this.context, this, "ShadingColor", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "style", {
			get: function () {
				_throwIfNotLoaded("style", this._St, _typeTable, this._isNull);
				return this._St;
			},
			set: function (value) {
				this._St=value;
				_createSetPropertyAction(this.context, this, "Style", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "styleBandedColumns", {
			get: function () {
				_throwIfNotLoaded("styleBandedColumns", this._Sty, _typeTable, this._isNull);
				return this._Sty;
			},
			set: function (value) {
				this._Sty=value;
				_createSetPropertyAction(this.context, this, "StyleBandedColumns", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "styleBandedRows", {
			get: function () {
				_throwIfNotLoaded("styleBandedRows", this._Styl, _typeTable, this._isNull);
				return this._Styl;
			},
			set: function (value) {
				this._Styl=value;
				_createSetPropertyAction(this.context, this, "StyleBandedRows", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "styleBuiltIn", {
			get: function () {
				_throwIfNotLoaded("styleBuiltIn", this._Style, _typeTable, this._isNull);
				return this._Style;
			},
			set: function (value) {
				this._Style=value;
				_createSetPropertyAction(this.context, this, "StyleBuiltIn", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "styleFirstColumn", {
			get: function () {
				_throwIfNotLoaded("styleFirstColumn", this._StyleF, _typeTable, this._isNull);
				return this._StyleF;
			},
			set: function (value) {
				this._StyleF=value;
				_createSetPropertyAction(this.context, this, "StyleFirstColumn", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "styleLastColumn", {
			get: function () {
				_throwIfNotLoaded("styleLastColumn", this._StyleL, _typeTable, this._isNull);
				return this._StyleL;
			},
			set: function (value) {
				this._StyleL=value;
				_createSetPropertyAction(this.context, this, "StyleLastColumn", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "styleTotalRow", {
			get: function () {
				_throwIfNotLoaded("styleTotalRow", this._StyleT, _typeTable, this._isNull);
				return this._StyleT;
			},
			set: function (value) {
				this._StyleT=value;
				_createSetPropertyAction(this.context, this, "StyleTotalRow", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "values", {
			get: function () {
				_throwIfNotLoaded("values", this._V, _typeTable, this._isNull);
				return this._V;
			},
			set: function (value) {
				this._V=value;
				_createSetPropertyAction(this.context, this, "Values", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "verticalAlignment", {
			get: function () {
				_throwIfNotLoaded("verticalAlignment", this._Ve, _typeTable, this._isNull);
				return this._Ve;
			},
			set: function (value) {
				this._Ve=value;
				_createSetPropertyAction(this.context, this, "VerticalAlignment", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "width", {
			get: function () {
				_throwIfNotLoaded("width", this._W, _typeTable, this._isNull);
				return this._W;
			},
			set: function (value) {
				this._W=value;
				_createSetPropertyAction(this.context, this, "Width", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "_Id", {
			get: function () {
				_throwIfNotLoaded("_Id", this.__I, _typeTable, this._isNull);
				return this.__I;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(Table.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.__R, _typeTable, this._isNull);
				return this.__R;
			},
			enumerable: true,
			configurable: true
		});
		Table.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["values", "style", "headerRowCount", "styleTotalRow", "styleFirstColumn", "styleLastColumn", "styleBandedRows", "styleBandedColumns", "shadingColor", "horizontalAlignment", "verticalAlignment", "width", "styleBuiltIn", "alignment"], ["font"], [
				"parentBody",
				"parentContentControl",
				"parentContentControlOrNullObject",
				"parentTable",
				"parentTableCell",
				"parentTableCellOrNullObject",
				"parentTableOrNullObject",
				"rows",
				"tables"
			]);
		};
		Table.prototype.update=function (properties) {
			this._recursivelyUpdate(properties);
		};
		Table.prototype.addColumns=function (insertLocation, columnCount, values) {
			_createMethodAction(this.context, this, "AddColumns", 0, [insertLocation, columnCount, values], false);
		};
		Table.prototype.addRows=function (insertLocation, rowCount, values) {
			return new Word.TableRowCollection(this.context, _createMethodObjectPath(this.context, this, "AddRows", 0, [insertLocation, rowCount, values], true, false, null, false));
		};
		Table.prototype.autoFitWindow=function () {
			_createMethodAction(this.context, this, "AutoFitWindow", 0, [], false);
		};
		Table.prototype.clear=function () {
			_createMethodAction(this.context, this, "Clear", 0, [], false);
		};
		Table.prototype.delete=function () {
			_createMethodAction(this.context, this, "Delete", 0, [], false);
		};
		Table.prototype.deleteColumns=function (columnIndex, columnCount) {
			_createMethodAction(this.context, this, "DeleteColumns", 0, [columnIndex, columnCount], false);
		};
		Table.prototype.deleteRows=function (rowIndex, rowCount) {
			_createMethodAction(this.context, this, "DeleteRows", 0, [rowIndex, rowCount], false);
		};
		Table.prototype.distributeColumns=function () {
			_createMethodAction(this.context, this, "DistributeColumns", 0, [], false);
		};
		Table.prototype.getBorder=function (borderLocation) {
			return new Word.TableBorder(this.context, _createMethodObjectPath(this.context, this, "GetBorder", 1, [borderLocation], false, false, null, false));
		};
		Table.prototype.getCell=function (rowIndex, cellIndex) {
			return new Word.TableCell(this.context, _createMethodObjectPath(this.context, this, "GetCell", 1, [rowIndex, cellIndex], false, false, null, false));
		};
		Table.prototype.getCellOrNullObject=function (rowIndex, cellIndex) {
			return new Word.TableCell(this.context, _createMethodObjectPath(this.context, this, "GetCellOrNullObject", 1, [rowIndex, cellIndex], false, false, null, false));
		};
		Table.prototype.getCellPadding=function (cellPaddingLocation) {
			var action=_createMethodAction(this.context, this, "GetCellPadding", 1, [cellPaddingLocation], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		Table.prototype.getNext=function () {
			return new Word.Table(this.context, _createMethodObjectPath(this.context, this, "GetNext", 1, [], false, false, null, false));
		};
		Table.prototype.getNextOrNullObject=function () {
			return new Word.Table(this.context, _createMethodObjectPath(this.context, this, "GetNextOrNullObject", 1, [], false, false, null, false));
		};
		Table.prototype.getParagraphAfter=function () {
			return new Word.Paragraph(this.context, _createMethodObjectPath(this.context, this, "GetParagraphAfter", 1, [], false, false, null, false));
		};
		Table.prototype.getParagraphAfterOrNullObject=function () {
			return new Word.Paragraph(this.context, _createMethodObjectPath(this.context, this, "GetParagraphAfterOrNullObject", 1, [], false, false, null, false));
		};
		Table.prototype.getParagraphBefore=function () {
			return new Word.Paragraph(this.context, _createMethodObjectPath(this.context, this, "GetParagraphBefore", 1, [], false, false, null, false));
		};
		Table.prototype.getParagraphBeforeOrNullObject=function () {
			return new Word.Paragraph(this.context, _createMethodObjectPath(this.context, this, "GetParagraphBeforeOrNullObject", 1, [], false, false, null, false));
		};
		Table.prototype.getRange=function (rangeLocation) {
			return new Word.Range(this.context, _createMethodObjectPath(this.context, this, "GetRange", 1, [rangeLocation], false, false, null, false));
		};
		Table.prototype.insertContentControl=function () {
			return new Word.ContentControl(this.context, _createMethodObjectPath(this.context, this, "InsertContentControl", 0, [], false, true, null, false));
		};
		Table.prototype.insertParagraph=function (paragraphText, insertLocation) {
			return new Word.Paragraph(this.context, _createMethodObjectPath(this.context, this, "InsertParagraph", 0, [paragraphText, insertLocation], false, true, null, false));
		};
		Table.prototype.insertTable=function (rowCount, columnCount, insertLocation, values) {
			return new Word.Table(this.context, _createMethodObjectPath(this.context, this, "InsertTable", 0, [rowCount, columnCount, insertLocation, values], false, true, null, false));
		};
		Table.prototype.search=function (searchText, searchOptions) {
			searchOptions=_normalizeSearchOptions(this.context, searchOptions);
			return new Word.RangeCollection(this.context, _createMethodObjectPath(this.context, this, "Search", 1, [searchText, searchOptions], true, false, null, false));
		};
		Table.prototype.select=function (selectionMode) {
			_createMethodAction(this.context, this, "Select", 1, [selectionMode], false);
		};
		Table.prototype.setCellPadding=function (cellPaddingLocation, cellPadding) {
			_createMethodAction(this.context, this, "SetCellPadding", 0, [cellPaddingLocation, cellPadding], false);
		};
		Table.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1, [], false);
		};
		Table.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Alignment"])) {
				this._A=obj["Alignment"];
			}
			if (!_isUndefined(obj["HeaderRowCount"])) {
				this._H=obj["HeaderRowCount"];
			}
			if (!_isUndefined(obj["HorizontalAlignment"])) {
				this._Ho=obj["HorizontalAlignment"];
			}
			if (!_isUndefined(obj["IsUniform"])) {
				this._I=obj["IsUniform"];
			}
			if (!_isUndefined(obj["NestingLevel"])) {
				this._N=obj["NestingLevel"];
			}
			if (!_isUndefined(obj["RowCount"])) {
				this._R=obj["RowCount"];
			}
			if (!_isUndefined(obj["ShadingColor"])) {
				this._S=obj["ShadingColor"];
			}
			if (!_isUndefined(obj["Style"])) {
				this._St=obj["Style"];
			}
			if (!_isUndefined(obj["StyleBandedColumns"])) {
				this._Sty=obj["StyleBandedColumns"];
			}
			if (!_isUndefined(obj["StyleBandedRows"])) {
				this._Styl=obj["StyleBandedRows"];
			}
			if (!_isUndefined(obj["StyleBuiltIn"])) {
				this._Style=obj["StyleBuiltIn"];
			}
			if (!_isUndefined(obj["StyleFirstColumn"])) {
				this._StyleF=obj["StyleFirstColumn"];
			}
			if (!_isUndefined(obj["StyleLastColumn"])) {
				this._StyleL=obj["StyleLastColumn"];
			}
			if (!_isUndefined(obj["StyleTotalRow"])) {
				this._StyleT=obj["StyleTotalRow"];
			}
			if (!_isUndefined(obj["Values"])) {
				this._V=obj["Values"];
			}
			if (!_isUndefined(obj["VerticalAlignment"])) {
				this._Ve=obj["VerticalAlignment"];
			}
			if (!_isUndefined(obj["Width"])) {
				this._W=obj["Width"];
			}
			if (!_isUndefined(obj["_Id"])) {
				this.__I=obj["_Id"];
			}
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.__R=obj["_ReferenceId"];
			}
			_handleNavigationPropertyResults(this, obj, ["font", "Font", "parentBody", "ParentBody", "parentContentControl", "ParentContentControl", "parentContentControlOrNullObject", "ParentContentControlOrNullObject", "parentTable", "ParentTable", "parentTableCell", "ParentTableCell", "parentTableCellOrNullObject", "ParentTableCellOrNullObject", "parentTableOrNullObject", "ParentTableOrNullObject", "rows", "Rows", "tables", "Tables"]);
		};
		Table.prototype.load=function (option) {
			return _load(this, option);
		};
		Table.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		Table.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["_ReferenceId"])) {
				this.__R=value["_ReferenceId"];
			}
			if (!_isUndefined(value["_Id"])) {
				this.__I=value["_Id"];
			}
		};
		Table.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result);
		};
		Table.prototype.track=function () {
			this.context.trackedObjects.add(this);
			return this;
		};
		Table.prototype.untrack=function () {
			this.context.trackedObjects.remove(this);
			return this;
		};
		Table.prototype.toJSON=function () {
			return _toJson(this, {
				"alignment": this._A,
				"headerRowCount": this._H,
				"horizontalAlignment": this._Ho,
				"isUniform": this._I,
				"nestingLevel": this._N,
				"rowCount": this._R,
				"shadingColor": this._S,
				"style": this._St,
				"styleBandedColumns": this._Sty,
				"styleBandedRows": this._Styl,
				"styleBuiltIn": this._Style,
				"styleFirstColumn": this._StyleF,
				"styleLastColumn": this._StyleL,
				"styleTotalRow": this._StyleT,
				"values": this._V,
				"verticalAlignment": this._Ve,
				"width": this._W,
			}, {
				"font": this._F,
				"rows": this._Ro,
				"tables": this._T,
			});
		};
		Table.prototype.ensureUnchanged=function (data) {
			_createEnsureUnchangedAction(this.context, this, data);
			return;
		};
		return Table;
	}(OfficeExtension.ClientObject));
	Word.Table=Table;
	var _typeTableCollection="TableCollection";
	var TableCollection=(function (_super) {
		__extends(TableCollection, _super);
		function TableCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(TableCollection.prototype, "_className", {
			get: function () {
				return "TableCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCollection.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["_ReferenceId"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, _typeTableCollection, this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCollection.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.__R, _typeTableCollection, this._isNull);
				return this.__R;
			},
			enumerable: true,
			configurable: true
		});
		TableCollection.prototype.getFirst=function () {
			return new Word.Table(this.context, _createMethodObjectPath(this.context, this, "GetFirst", 1, [], false, false, null, false));
		};
		TableCollection.prototype.getFirstOrNullObject=function () {
			return new Word.Table(this.context, _createMethodObjectPath(this.context, this, "GetFirstOrNullObject", 1, [], false, false, null, false));
		};
		TableCollection.prototype._GetItem=function (index) {
			return new Word.Table(this.context, _createIndexerObjectPath(this.context, this, [index]));
		};
		TableCollection.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1, [], false);
		};
		TableCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.__R=obj["_ReferenceId"];
			}
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Word.Table(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		TableCollection.prototype.load=function (option) {
			return _load(this, option);
		};
		TableCollection.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		TableCollection.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["_ReferenceId"])) {
				this.__R=value["_ReferenceId"];
			}
		};
		TableCollection.prototype._handleRetrieveResult=function (value, result) {
			var _this=this;
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result, function (childItemData, index) { return new Word.Table(_this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, _this.context, _this, childItemData, index)); });
		};
		TableCollection.prototype.track=function () {
			this.context.trackedObjects.add(this);
			return this;
		};
		TableCollection.prototype.untrack=function () {
			this.context.trackedObjects.remove(this);
			return this;
		};
		TableCollection.prototype.toJSON=function () {
			return _toJson(this, {}, {}, this.m__items);
		};
		return TableCollection;
	}(OfficeExtension.ClientObject));
	Word.TableCollection=TableCollection;
	var _typeTableRow="TableRow";
	var TableRow=(function (_super) {
		__extends(TableRow, _super);
		function TableRow() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(TableRow.prototype, "_className", {
			get: function () {
				return "TableRow";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRow.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["_Id", "_ReferenceId", "cellCount", "rowIndex", "values", "shadingColor", "horizontalAlignment", "verticalAlignment", "isHeader", "preferredHeight"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRow.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, false, false, false, true, true, true, true, false, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRow.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["cells", "parentTable", "font"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRow.prototype, "cells", {
			get: function () {
				if (!this._Ce) {
					this._Ce=new Word.TableCellCollection(this.context, _createPropertyObjectPath(this.context, this, "Cells", true, false, false));
				}
				return this._Ce;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRow.prototype, "font", {
			get: function () {
				if (!this._F) {
					this._F=new Word.Font(this.context, _createPropertyObjectPath(this.context, this, "Font", false, false, false));
				}
				return this._F;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRow.prototype, "parentTable", {
			get: function () {
				if (!this._P) {
					this._P=new Word.Table(this.context, _createPropertyObjectPath(this.context, this, "ParentTable", false, false, false));
				}
				return this._P;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRow.prototype, "cellCount", {
			get: function () {
				_throwIfNotLoaded("cellCount", this._C, _typeTableRow, this._isNull);
				return this._C;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRow.prototype, "horizontalAlignment", {
			get: function () {
				_throwIfNotLoaded("horizontalAlignment", this._H, _typeTableRow, this._isNull);
				return this._H;
			},
			set: function (value) {
				this._H=value;
				_createSetPropertyAction(this.context, this, "HorizontalAlignment", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRow.prototype, "isHeader", {
			get: function () {
				_throwIfNotLoaded("isHeader", this._I, _typeTableRow, this._isNull);
				return this._I;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRow.prototype, "preferredHeight", {
			get: function () {
				_throwIfNotLoaded("preferredHeight", this._Pr, _typeTableRow, this._isNull);
				return this._Pr;
			},
			set: function (value) {
				this._Pr=value;
				_createSetPropertyAction(this.context, this, "PreferredHeight", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRow.prototype, "rowIndex", {
			get: function () {
				_throwIfNotLoaded("rowIndex", this._R, _typeTableRow, this._isNull);
				return this._R;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRow.prototype, "shadingColor", {
			get: function () {
				_throwIfNotLoaded("shadingColor", this._S, _typeTableRow, this._isNull);
				return this._S;
			},
			set: function (value) {
				this._S=value;
				_createSetPropertyAction(this.context, this, "ShadingColor", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRow.prototype, "values", {
			get: function () {
				_throwIfNotLoaded("values", this._V, _typeTableRow, this._isNull);
				return this._V;
			},
			set: function (value) {
				this._V=value;
				_createSetPropertyAction(this.context, this, "Values", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRow.prototype, "verticalAlignment", {
			get: function () {
				_throwIfNotLoaded("verticalAlignment", this._Ve, _typeTableRow, this._isNull);
				return this._Ve;
			},
			set: function (value) {
				this._Ve=value;
				_createSetPropertyAction(this.context, this, "VerticalAlignment", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRow.prototype, "_Id", {
			get: function () {
				_throwIfNotLoaded("_Id", this.__I, _typeTableRow, this._isNull);
				return this.__I;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRow.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.__R, _typeTableRow, this._isNull);
				return this.__R;
			},
			enumerable: true,
			configurable: true
		});
		TableRow.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["values", "shadingColor", "horizontalAlignment", "verticalAlignment", "preferredHeight"], ["font"], [
				"cells",
				"parentTable"
			]);
		};
		TableRow.prototype.update=function (properties) {
			this._recursivelyUpdate(properties);
		};
		TableRow.prototype.clear=function () {
			_createMethodAction(this.context, this, "Clear", 0, [], false);
		};
		TableRow.prototype.delete=function () {
			_createMethodAction(this.context, this, "Delete", 0, [], false);
		};
		TableRow.prototype.getBorder=function (borderLocation) {
			return new Word.TableBorder(this.context, _createMethodObjectPath(this.context, this, "GetBorder", 1, [borderLocation], false, false, null, false));
		};
		TableRow.prototype.getCellPadding=function (cellPaddingLocation) {
			var action=_createMethodAction(this.context, this, "GetCellPadding", 1, [cellPaddingLocation], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		TableRow.prototype.getNext=function () {
			return new Word.TableRow(this.context, _createMethodObjectPath(this.context, this, "GetNext", 1, [], false, false, null, false));
		};
		TableRow.prototype.getNextOrNullObject=function () {
			return new Word.TableRow(this.context, _createMethodObjectPath(this.context, this, "GetNextOrNullObject", 1, [], false, false, null, false));
		};
		TableRow.prototype.insertRows=function (insertLocation, rowCount, values) {
			return new Word.TableRowCollection(this.context, _createMethodObjectPath(this.context, this, "InsertRows", 1, [insertLocation, rowCount, values], true, false, null, false));
		};
		TableRow.prototype.search=function (searchText, searchOptions) {
			searchOptions=_normalizeSearchOptions(this.context, searchOptions);
			return new Word.RangeCollection(this.context, _createMethodObjectPath(this.context, this, "Search", 1, [searchText, searchOptions], true, false, null, false));
		};
		TableRow.prototype.select=function (selectionMode) {
			_createMethodAction(this.context, this, "Select", 1, [selectionMode], false);
		};
		TableRow.prototype.setCellPadding=function (cellPaddingLocation, cellPadding) {
			_createMethodAction(this.context, this, "SetCellPadding", 0, [cellPaddingLocation, cellPadding], false);
		};
		TableRow.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1, [], false);
		};
		TableRow.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["CellCount"])) {
				this._C=obj["CellCount"];
			}
			if (!_isUndefined(obj["HorizontalAlignment"])) {
				this._H=obj["HorizontalAlignment"];
			}
			if (!_isUndefined(obj["IsHeader"])) {
				this._I=obj["IsHeader"];
			}
			if (!_isUndefined(obj["PreferredHeight"])) {
				this._Pr=obj["PreferredHeight"];
			}
			if (!_isUndefined(obj["RowIndex"])) {
				this._R=obj["RowIndex"];
			}
			if (!_isUndefined(obj["ShadingColor"])) {
				this._S=obj["ShadingColor"];
			}
			if (!_isUndefined(obj["Values"])) {
				this._V=obj["Values"];
			}
			if (!_isUndefined(obj["VerticalAlignment"])) {
				this._Ve=obj["VerticalAlignment"];
			}
			if (!_isUndefined(obj["_Id"])) {
				this.__I=obj["_Id"];
			}
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.__R=obj["_ReferenceId"];
			}
			_handleNavigationPropertyResults(this, obj, ["cells", "Cells", "font", "Font", "parentTable", "ParentTable"]);
		};
		TableRow.prototype.load=function (option) {
			return _load(this, option);
		};
		TableRow.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		TableRow.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["_ReferenceId"])) {
				this.__R=value["_ReferenceId"];
			}
			if (!_isUndefined(value["_Id"])) {
				this.__I=value["_Id"];
			}
		};
		TableRow.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result);
		};
		TableRow.prototype.track=function () {
			this.context.trackedObjects.add(this);
			return this;
		};
		TableRow.prototype.untrack=function () {
			this.context.trackedObjects.remove(this);
			return this;
		};
		TableRow.prototype.toJSON=function () {
			return _toJson(this, {
				"cellCount": this._C,
				"horizontalAlignment": this._H,
				"isHeader": this._I,
				"preferredHeight": this._Pr,
				"rowIndex": this._R,
				"shadingColor": this._S,
				"values": this._V,
				"verticalAlignment": this._Ve,
			}, {
				"cells": this._Ce,
				"font": this._F,
			});
		};
		TableRow.prototype.ensureUnchanged=function (data) {
			_createEnsureUnchangedAction(this.context, this, data);
			return;
		};
		return TableRow;
	}(OfficeExtension.ClientObject));
	Word.TableRow=TableRow;
	var _typeTableRowCollection="TableRowCollection";
	var TableRowCollection=(function (_super) {
		__extends(TableRowCollection, _super);
		function TableRowCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(TableRowCollection.prototype, "_className", {
			get: function () {
				return "TableRowCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRowCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRowCollection.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["_ReferenceId"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRowCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, _typeTableRowCollection, this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableRowCollection.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.__R, _typeTableRowCollection, this._isNull);
				return this.__R;
			},
			enumerable: true,
			configurable: true
		});
		TableRowCollection.prototype.getFirst=function () {
			return new Word.TableRow(this.context, _createMethodObjectPath(this.context, this, "GetFirst", 1, [], false, false, null, false));
		};
		TableRowCollection.prototype.getFirstOrNullObject=function () {
			return new Word.TableRow(this.context, _createMethodObjectPath(this.context, this, "GetFirstOrNullObject", 1, [], false, false, null, false));
		};
		TableRowCollection.prototype._GetItem=function (index) {
			return new Word.TableRow(this.context, _createIndexerObjectPath(this.context, this, [index]));
		};
		TableRowCollection.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1, [], false);
		};
		TableRowCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.__R=obj["_ReferenceId"];
			}
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Word.TableRow(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		TableRowCollection.prototype.load=function (option) {
			return _load(this, option);
		};
		TableRowCollection.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		TableRowCollection.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["_ReferenceId"])) {
				this.__R=value["_ReferenceId"];
			}
		};
		TableRowCollection.prototype._handleRetrieveResult=function (value, result) {
			var _this=this;
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result, function (childItemData, index) { return new Word.TableRow(_this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, _this.context, _this, childItemData, index)); });
		};
		TableRowCollection.prototype.track=function () {
			this.context.trackedObjects.add(this);
			return this;
		};
		TableRowCollection.prototype.untrack=function () {
			this.context.trackedObjects.remove(this);
			return this;
		};
		TableRowCollection.prototype.toJSON=function () {
			return _toJson(this, {}, {}, this.m__items);
		};
		return TableRowCollection;
	}(OfficeExtension.ClientObject));
	Word.TableRowCollection=TableRowCollection;
	var _typeTableCell="TableCell";
	var TableCell=(function (_super) {
		__extends(TableCell, _super);
		function TableCell() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(TableCell.prototype, "_className", {
			get: function () {
				return "TableCell";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCell.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["_Id", "_ReferenceId", "rowIndex", "cellIndex", "value", "shadingColor", "horizontalAlignment", "verticalAlignment", "columnWidth", "width"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCell.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, false, false, false, true, true, true, true, true, false];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCell.prototype, "_navigationPropertyNames", {
			get: function () {
				return ["parentTable", "parentRow", "body"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCell.prototype, "body", {
			get: function () {
				if (!this._B) {
					this._B=new Word.Body(this.context, _createPropertyObjectPath(this.context, this, "Body", false, false, false));
				}
				return this._B;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCell.prototype, "parentRow", {
			get: function () {
				if (!this._P) {
					this._P=new Word.TableRow(this.context, _createPropertyObjectPath(this.context, this, "ParentRow", false, false, false));
				}
				return this._P;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCell.prototype, "parentTable", {
			get: function () {
				if (!this._Pa) {
					this._Pa=new Word.Table(this.context, _createPropertyObjectPath(this.context, this, "ParentTable", false, false, false));
				}
				return this._Pa;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCell.prototype, "cellIndex", {
			get: function () {
				_throwIfNotLoaded("cellIndex", this._C, _typeTableCell, this._isNull);
				return this._C;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCell.prototype, "columnWidth", {
			get: function () {
				_throwIfNotLoaded("columnWidth", this._Co, _typeTableCell, this._isNull);
				return this._Co;
			},
			set: function (value) {
				this._Co=value;
				_createSetPropertyAction(this.context, this, "ColumnWidth", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCell.prototype, "horizontalAlignment", {
			get: function () {
				_throwIfNotLoaded("horizontalAlignment", this._H, _typeTableCell, this._isNull);
				return this._H;
			},
			set: function (value) {
				this._H=value;
				_createSetPropertyAction(this.context, this, "HorizontalAlignment", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCell.prototype, "rowIndex", {
			get: function () {
				_throwIfNotLoaded("rowIndex", this._R, _typeTableCell, this._isNull);
				return this._R;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCell.prototype, "shadingColor", {
			get: function () {
				_throwIfNotLoaded("shadingColor", this._S, _typeTableCell, this._isNull);
				return this._S;
			},
			set: function (value) {
				this._S=value;
				_createSetPropertyAction(this.context, this, "ShadingColor", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCell.prototype, "value", {
			get: function () {
				_throwIfNotLoaded("value", this._V, _typeTableCell, this._isNull);
				return this._V;
			},
			set: function (value) {
				this._V=value;
				_createSetPropertyAction(this.context, this, "Value", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCell.prototype, "verticalAlignment", {
			get: function () {
				_throwIfNotLoaded("verticalAlignment", this._Ve, _typeTableCell, this._isNull);
				return this._Ve;
			},
			set: function (value) {
				this._Ve=value;
				_createSetPropertyAction(this.context, this, "VerticalAlignment", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCell.prototype, "width", {
			get: function () {
				_throwIfNotLoaded("width", this._W, _typeTableCell, this._isNull);
				return this._W;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCell.prototype, "_Id", {
			get: function () {
				_throwIfNotLoaded("_Id", this.__I, _typeTableCell, this._isNull);
				return this.__I;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCell.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.__R, _typeTableCell, this._isNull);
				return this.__R;
			},
			enumerable: true,
			configurable: true
		});
		TableCell.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["value", "shadingColor", "horizontalAlignment", "verticalAlignment", "columnWidth"], ["body"], [
				"parentRow",
				"parentTable"
			]);
		};
		TableCell.prototype.update=function (properties) {
			this._recursivelyUpdate(properties);
		};
		TableCell.prototype.deleteColumn=function () {
			_createMethodAction(this.context, this, "DeleteColumn", 0, [], false);
		};
		TableCell.prototype.deleteRow=function () {
			_createMethodAction(this.context, this, "DeleteRow", 0, [], false);
		};
		TableCell.prototype.getBorder=function (borderLocation) {
			return new Word.TableBorder(this.context, _createMethodObjectPath(this.context, this, "GetBorder", 1, [borderLocation], false, false, null, false));
		};
		TableCell.prototype.getCellPadding=function (cellPaddingLocation) {
			var action=_createMethodAction(this.context, this, "GetCellPadding", 1, [cellPaddingLocation], false);
			var ret=new OfficeExtension.ClientResult();
			_addActionResultHandler(this, action, ret);
			return ret;
		};
		TableCell.prototype.getNext=function () {
			return new Word.TableCell(this.context, _createMethodObjectPath(this.context, this, "GetNext", 1, [], false, false, null, false));
		};
		TableCell.prototype.getNextOrNullObject=function () {
			return new Word.TableCell(this.context, _createMethodObjectPath(this.context, this, "GetNextOrNullObject", 1, [], false, false, null, false));
		};
		TableCell.prototype.insertColumns=function (insertLocation, columnCount, values) {
			_createMethodAction(this.context, this, "InsertColumns", 0, [insertLocation, columnCount, values], false);
		};
		TableCell.prototype.insertRows=function (insertLocation, rowCount, values) {
			return new Word.TableRowCollection(this.context, _createMethodObjectPath(this.context, this, "InsertRows", 0, [insertLocation, rowCount, values], true, false, null, false));
		};
		TableCell.prototype.setCellPadding=function (cellPaddingLocation, cellPadding) {
			_createMethodAction(this.context, this, "SetCellPadding", 0, [cellPaddingLocation, cellPadding], false);
		};
		TableCell.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1, [], false);
		};
		TableCell.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["CellIndex"])) {
				this._C=obj["CellIndex"];
			}
			if (!_isUndefined(obj["ColumnWidth"])) {
				this._Co=obj["ColumnWidth"];
			}
			if (!_isUndefined(obj["HorizontalAlignment"])) {
				this._H=obj["HorizontalAlignment"];
			}
			if (!_isUndefined(obj["RowIndex"])) {
				this._R=obj["RowIndex"];
			}
			if (!_isUndefined(obj["ShadingColor"])) {
				this._S=obj["ShadingColor"];
			}
			if (!_isUndefined(obj["Value"])) {
				this._V=obj["Value"];
			}
			if (!_isUndefined(obj["VerticalAlignment"])) {
				this._Ve=obj["VerticalAlignment"];
			}
			if (!_isUndefined(obj["Width"])) {
				this._W=obj["Width"];
			}
			if (!_isUndefined(obj["_Id"])) {
				this.__I=obj["_Id"];
			}
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.__R=obj["_ReferenceId"];
			}
			_handleNavigationPropertyResults(this, obj, ["body", "Body", "parentRow", "ParentRow", "parentTable", "ParentTable"]);
		};
		TableCell.prototype.load=function (option) {
			return _load(this, option);
		};
		TableCell.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		TableCell.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["_ReferenceId"])) {
				this.__R=value["_ReferenceId"];
			}
			if (!_isUndefined(value["_Id"])) {
				this.__I=value["_Id"];
			}
		};
		TableCell.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result);
		};
		TableCell.prototype.track=function () {
			this.context.trackedObjects.add(this);
			return this;
		};
		TableCell.prototype.untrack=function () {
			this.context.trackedObjects.remove(this);
			return this;
		};
		TableCell.prototype.toJSON=function () {
			return _toJson(this, {
				"cellIndex": this._C,
				"columnWidth": this._Co,
				"horizontalAlignment": this._H,
				"rowIndex": this._R,
				"shadingColor": this._S,
				"value": this._V,
				"verticalAlignment": this._Ve,
				"width": this._W,
			}, {
				"body": this._B,
			});
		};
		TableCell.prototype.ensureUnchanged=function (data) {
			_createEnsureUnchangedAction(this.context, this, data);
			return;
		};
		return TableCell;
	}(OfficeExtension.ClientObject));
	Word.TableCell=TableCell;
	var _typeTableCellCollection="TableCellCollection";
	var TableCellCollection=(function (_super) {
		__extends(TableCellCollection, _super);
		function TableCellCollection() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(TableCellCollection.prototype, "_className", {
			get: function () {
				return "TableCellCollection";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCellCollection.prototype, "_isCollection", {
			get: function () {
				return true;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCellCollection.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["_ReferenceId"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCellCollection.prototype, "items", {
			get: function () {
				_throwIfNotLoaded("items", this.m__items, _typeTableCellCollection, this._isNull);
				return this.m__items;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableCellCollection.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.__R, _typeTableCellCollection, this._isNull);
				return this.__R;
			},
			enumerable: true,
			configurable: true
		});
		TableCellCollection.prototype.getFirst=function () {
			return new Word.TableCell(this.context, _createMethodObjectPath(this.context, this, "GetFirst", 1, [], false, false, null, false));
		};
		TableCellCollection.prototype.getFirstOrNullObject=function () {
			return new Word.TableCell(this.context, _createMethodObjectPath(this.context, this, "GetFirstOrNullObject", 1, [], false, false, null, false));
		};
		TableCellCollection.prototype._GetItem=function (index) {
			return new Word.TableCell(this.context, _createIndexerObjectPath(this.context, this, [index]));
		};
		TableCellCollection.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1, [], false);
		};
		TableCellCollection.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.__R=obj["_ReferenceId"];
			}
			if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
				this.m__items=[];
				var _data=obj[OfficeExtension.Constants.items];
				for (var i=0; i < _data.length; i++) {
					var _item=new Word.TableCell(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
					_item._handleResult(_data[i]);
					this.m__items.push(_item);
				}
			}
		};
		TableCellCollection.prototype.load=function (option) {
			return _load(this, option);
		};
		TableCellCollection.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		TableCellCollection.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["_ReferenceId"])) {
				this.__R=value["_ReferenceId"];
			}
		};
		TableCellCollection.prototype._handleRetrieveResult=function (value, result) {
			var _this=this;
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result, function (childItemData, index) { return new Word.TableCell(_this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, _this.context, _this, childItemData, index)); });
		};
		TableCellCollection.prototype.track=function () {
			this.context.trackedObjects.add(this);
			return this;
		};
		TableCellCollection.prototype.untrack=function () {
			this.context.trackedObjects.remove(this);
			return this;
		};
		TableCellCollection.prototype.toJSON=function () {
			return _toJson(this, {}, {}, this.m__items);
		};
		return TableCellCollection;
	}(OfficeExtension.ClientObject));
	Word.TableCellCollection=TableCellCollection;
	var _typeTableBorder="TableBorder";
	var TableBorder=(function (_super) {
		__extends(TableBorder, _super);
		function TableBorder() {
			return _super !==null && _super.apply(this, arguments) || this;
		}
		Object.defineProperty(TableBorder.prototype, "_className", {
			get: function () {
				return "TableBorder";
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableBorder.prototype, "_scalarPropertyNames", {
			get: function () {
				return ["_ReferenceId", "color", "type", "width"];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableBorder.prototype, "_scalarPropertyUpdateable", {
			get: function () {
				return [false, true, true, true];
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableBorder.prototype, "color", {
			get: function () {
				_throwIfNotLoaded("color", this._C, _typeTableBorder, this._isNull);
				return this._C;
			},
			set: function (value) {
				this._C=value;
				_createSetPropertyAction(this.context, this, "Color", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableBorder.prototype, "type", {
			get: function () {
				_throwIfNotLoaded("type", this._T, _typeTableBorder, this._isNull);
				return this._T;
			},
			set: function (value) {
				this._T=value;
				_createSetPropertyAction(this.context, this, "Type", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableBorder.prototype, "width", {
			get: function () {
				_throwIfNotLoaded("width", this._W, _typeTableBorder, this._isNull);
				return this._W;
			},
			set: function (value) {
				this._W=value;
				_createSetPropertyAction(this.context, this, "Width", value);
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(TableBorder.prototype, "_ReferenceId", {
			get: function () {
				_throwIfNotLoaded("_ReferenceId", this.__R, _typeTableBorder, this._isNull);
				return this.__R;
			},
			enumerable: true,
			configurable: true
		});
		TableBorder.prototype.set=function (properties, options) {
			this._recursivelySet(properties, options, ["color", "type", "width"], [], []);
		};
		TableBorder.prototype.update=function (properties) {
			this._recursivelyUpdate(properties);
		};
		TableBorder.prototype._KeepReference=function () {
			_createMethodAction(this.context, this, "_KeepReference", 1, [], false);
		};
		TableBorder.prototype._handleResult=function (value) {
			_super.prototype._handleResult.call(this, value);
			if (_isNullOrUndefined(value))
				return;
			var obj=value;
			_fixObjectPathIfNecessary(this, obj);
			if (!_isUndefined(obj["Color"])) {
				this._C=obj["Color"];
			}
			if (!_isUndefined(obj["Type"])) {
				this._T=obj["Type"];
			}
			if (!_isUndefined(obj["Width"])) {
				this._W=obj["Width"];
			}
			if (!_isUndefined(obj["_ReferenceId"])) {
				this.__R=obj["_ReferenceId"];
			}
		};
		TableBorder.prototype.load=function (option) {
			return _load(this, option);
		};
		TableBorder.prototype.retrieve=function (option) {
			return _retrieve(this, option);
		};
		TableBorder.prototype._handleIdResult=function (value) {
			_super.prototype._handleIdResult.call(this, value);
			if (_isNullOrUndefined(value)) {
				return;
			}
			if (!_isUndefined(value["_ReferenceId"])) {
				this.__R=value["_ReferenceId"];
			}
		};
		TableBorder.prototype._handleRetrieveResult=function (value, result) {
			_super.prototype._handleRetrieveResult.call(this, value, result);
			_processRetrieveResult(this, value, result);
		};
		TableBorder.prototype.track=function () {
			this.context.trackedObjects.add(this);
			return this;
		};
		TableBorder.prototype.untrack=function () {
			this.context.trackedObjects.remove(this);
			return this;
		};
		TableBorder.prototype.toJSON=function () {
			return _toJson(this, {
				"color": this._C,
				"type": this._T,
				"width": this._W,
			}, {});
		};
		TableBorder.prototype.ensureUnchanged=function (data) {
			_createEnsureUnchangedAction(this.context, this, data);
			return;
		};
		return TableBorder;
	}(OfficeExtension.ClientObject));
	Word.TableBorder=TableBorder;
	var ContentControlType;
	(function (ContentControlType) {
		ContentControlType["unknown"]="Unknown";
		ContentControlType["richTextInline"]="RichTextInline";
		ContentControlType["richTextParagraphs"]="RichTextParagraphs";
		ContentControlType["richTextTableCell"]="RichTextTableCell";
		ContentControlType["richTextTableRow"]="RichTextTableRow";
		ContentControlType["richTextTable"]="RichTextTable";
		ContentControlType["plainTextInline"]="PlainTextInline";
		ContentControlType["plainTextParagraph"]="PlainTextParagraph";
		ContentControlType["picture"]="Picture";
		ContentControlType["buildingBlockGallery"]="BuildingBlockGallery";
		ContentControlType["checkBox"]="CheckBox";
		ContentControlType["comboBox"]="ComboBox";
		ContentControlType["dropDownList"]="DropDownList";
		ContentControlType["datePicker"]="DatePicker";
		ContentControlType["repeatingSection"]="RepeatingSection";
		ContentControlType["richText"]="RichText";
		ContentControlType["plainText"]="PlainText";
	})(ContentControlType=Word.ContentControlType || (Word.ContentControlType={}));
	var ContentControlAppearance;
	(function (ContentControlAppearance) {
		ContentControlAppearance["boundingBox"]="BoundingBox";
		ContentControlAppearance["tags"]="Tags";
		ContentControlAppearance["hidden"]="Hidden";
	})(ContentControlAppearance=Word.ContentControlAppearance || (Word.ContentControlAppearance={}));
	var UnderlineType;
	(function (UnderlineType) {
		UnderlineType["mixed"]="Mixed";
		UnderlineType["none"]="None";
		UnderlineType["hidden"]="Hidden";
		UnderlineType["dotLine"]="DotLine";
		UnderlineType["single"]="Single";
		UnderlineType["word"]="Word";
		UnderlineType["double"]="Double";
		UnderlineType["thick"]="Thick";
		UnderlineType["dotted"]="Dotted";
		UnderlineType["dottedHeavy"]="DottedHeavy";
		UnderlineType["dashLine"]="DashLine";
		UnderlineType["dashLineHeavy"]="DashLineHeavy";
		UnderlineType["dashLineLong"]="DashLineLong";
		UnderlineType["dashLineLongHeavy"]="DashLineLongHeavy";
		UnderlineType["dotDashLine"]="DotDashLine";
		UnderlineType["dotDashLineHeavy"]="DotDashLineHeavy";
		UnderlineType["twoDotDashLine"]="TwoDotDashLine";
		UnderlineType["twoDotDashLineHeavy"]="TwoDotDashLineHeavy";
		UnderlineType["wave"]="Wave";
		UnderlineType["waveHeavy"]="WaveHeavy";
		UnderlineType["waveDouble"]="WaveDouble";
	})(UnderlineType=Word.UnderlineType || (Word.UnderlineType={}));
	var BreakType;
	(function (BreakType) {
		BreakType["page"]="Page";
		BreakType["next"]="Next";
		BreakType["sectionNext"]="SectionNext";
		BreakType["sectionContinuous"]="SectionContinuous";
		BreakType["sectionEven"]="SectionEven";
		BreakType["sectionOdd"]="SectionOdd";
		BreakType["line"]="Line";
	})(BreakType=Word.BreakType || (Word.BreakType={}));
	var InsertLocation;
	(function (InsertLocation) {
		InsertLocation["before"]="Before";
		InsertLocation["after"]="After";
		InsertLocation["start"]="Start";
		InsertLocation["end"]="End";
		InsertLocation["replace"]="Replace";
	})(InsertLocation=Word.InsertLocation || (Word.InsertLocation={}));
	var Alignment;
	(function (Alignment) {
		Alignment["mixed"]="Mixed";
		Alignment["unknown"]="Unknown";
		Alignment["left"]="Left";
		Alignment["centered"]="Centered";
		Alignment["right"]="Right";
		Alignment["justified"]="Justified";
	})(Alignment=Word.Alignment || (Word.Alignment={}));
	var HeaderFooterType;
	(function (HeaderFooterType) {
		HeaderFooterType["primary"]="Primary";
		HeaderFooterType["firstPage"]="FirstPage";
		HeaderFooterType["evenPages"]="EvenPages";
	})(HeaderFooterType=Word.HeaderFooterType || (Word.HeaderFooterType={}));
	var BodyType;
	(function (BodyType) {
		BodyType["unknown"]="Unknown";
		BodyType["mainDoc"]="MainDoc";
		BodyType["section"]="Section";
		BodyType["header"]="Header";
		BodyType["footer"]="Footer";
		BodyType["tableCell"]="TableCell";
	})(BodyType=Word.BodyType || (Word.BodyType={}));
	var SelectionMode;
	(function (SelectionMode) {
		SelectionMode["select"]="Select";
		SelectionMode["start"]="Start";
		SelectionMode["end"]="End";
	})(SelectionMode=Word.SelectionMode || (Word.SelectionMode={}));
	var ImageFormat;
	(function (ImageFormat) {
		ImageFormat["unsupported"]="Unsupported";
		ImageFormat["undefined"]="Undefined";
		ImageFormat["bmp"]="Bmp";
		ImageFormat["jpeg"]="Jpeg";
		ImageFormat["gif"]="Gif";
		ImageFormat["tiff"]="Tiff";
		ImageFormat["png"]="Png";
		ImageFormat["icon"]="Icon";
		ImageFormat["exif"]="Exif";
		ImageFormat["wmf"]="Wmf";
		ImageFormat["emf"]="Emf";
		ImageFormat["pict"]="Pict";
		ImageFormat["pdf"]="Pdf";
		ImageFormat["svg"]="Svg";
	})(ImageFormat=Word.ImageFormat || (Word.ImageFormat={}));
	var RangeLocation;
	(function (RangeLocation) {
		RangeLocation["whole"]="Whole";
		RangeLocation["start"]="Start";
		RangeLocation["end"]="End";
		RangeLocation["before"]="Before";
		RangeLocation["after"]="After";
		RangeLocation["content"]="Content";
	})(RangeLocation=Word.RangeLocation || (Word.RangeLocation={}));
	var LocationRelation;
	(function (LocationRelation) {
		LocationRelation["unrelated"]="Unrelated";
		LocationRelation["equal"]="Equal";
		LocationRelation["containsStart"]="ContainsStart";
		LocationRelation["containsEnd"]="ContainsEnd";
		LocationRelation["contains"]="Contains";
		LocationRelation["insideStart"]="InsideStart";
		LocationRelation["insideEnd"]="InsideEnd";
		LocationRelation["inside"]="Inside";
		LocationRelation["adjacentBefore"]="AdjacentBefore";
		LocationRelation["overlapsBefore"]="OverlapsBefore";
		LocationRelation["before"]="Before";
		LocationRelation["adjacentAfter"]="AdjacentAfter";
		LocationRelation["overlapsAfter"]="OverlapsAfter";
		LocationRelation["after"]="After";
	})(LocationRelation=Word.LocationRelation || (Word.LocationRelation={}));
	var BorderLocation;
	(function (BorderLocation) {
		BorderLocation["top"]="Top";
		BorderLocation["left"]="Left";
		BorderLocation["bottom"]="Bottom";
		BorderLocation["right"]="Right";
		BorderLocation["insideHorizontal"]="InsideHorizontal";
		BorderLocation["insideVertical"]="InsideVertical";
		BorderLocation["inside"]="Inside";
		BorderLocation["outside"]="Outside";
		BorderLocation["all"]="All";
	})(BorderLocation=Word.BorderLocation || (Word.BorderLocation={}));
	var CellPaddingLocation;
	(function (CellPaddingLocation) {
		CellPaddingLocation["top"]="Top";
		CellPaddingLocation["left"]="Left";
		CellPaddingLocation["bottom"]="Bottom";
		CellPaddingLocation["right"]="Right";
	})(CellPaddingLocation=Word.CellPaddingLocation || (Word.CellPaddingLocation={}));
	var BorderType;
	(function (BorderType) {
		BorderType["mixed"]="Mixed";
		BorderType["none"]="None";
		BorderType["single"]="Single";
		BorderType["double"]="Double";
		BorderType["dotted"]="Dotted";
		BorderType["dashed"]="Dashed";
		BorderType["dotDashed"]="DotDashed";
		BorderType["dot2Dashed"]="Dot2Dashed";
		BorderType["triple"]="Triple";
		BorderType["thinThickSmall"]="ThinThickSmall";
		BorderType["thickThinSmall"]="ThickThinSmall";
		BorderType["thinThickThinSmall"]="ThinThickThinSmall";
		BorderType["thinThickMed"]="ThinThickMed";
		BorderType["thickThinMed"]="ThickThinMed";
		BorderType["thinThickThinMed"]="ThinThickThinMed";
		BorderType["thinThickLarge"]="ThinThickLarge";
		BorderType["thickThinLarge"]="ThickThinLarge";
		BorderType["thinThickThinLarge"]="ThinThickThinLarge";
		BorderType["wave"]="Wave";
		BorderType["doubleWave"]="DoubleWave";
		BorderType["dashedSmall"]="DashedSmall";
		BorderType["dashDotStroked"]="DashDotStroked";
		BorderType["threeDEmboss"]="ThreeDEmboss";
		BorderType["threeDEngrave"]="ThreeDEngrave";
	})(BorderType=Word.BorderType || (Word.BorderType={}));
	var VerticalAlignment;
	(function (VerticalAlignment) {
		VerticalAlignment["mixed"]="Mixed";
		VerticalAlignment["top"]="Top";
		VerticalAlignment["center"]="Center";
		VerticalAlignment["bottom"]="Bottom";
	})(VerticalAlignment=Word.VerticalAlignment || (Word.VerticalAlignment={}));
	var ListLevelType;
	(function (ListLevelType) {
		ListLevelType["bullet"]="Bullet";
		ListLevelType["number"]="Number";
		ListLevelType["picture"]="Picture";
	})(ListLevelType=Word.ListLevelType || (Word.ListLevelType={}));
	var ListBullet;
	(function (ListBullet) {
		ListBullet["custom"]="Custom";
		ListBullet["solid"]="Solid";
		ListBullet["hollow"]="Hollow";
		ListBullet["square"]="Square";
		ListBullet["diamonds"]="Diamonds";
		ListBullet["arrow"]="Arrow";
		ListBullet["checkmark"]="Checkmark";
	})(ListBullet=Word.ListBullet || (Word.ListBullet={}));
	var ListNumbering;
	(function (ListNumbering) {
		ListNumbering["none"]="None";
		ListNumbering["arabic"]="Arabic";
		ListNumbering["upperRoman"]="UpperRoman";
		ListNumbering["lowerRoman"]="LowerRoman";
		ListNumbering["upperLetter"]="UpperLetter";
		ListNumbering["lowerLetter"]="LowerLetter";
	})(ListNumbering=Word.ListNumbering || (Word.ListNumbering={}));
	var Style;
	(function (Style) {
		Style["other"]="Other";
		Style["normal"]="Normal";
		Style["heading1"]="Heading1";
		Style["heading2"]="Heading2";
		Style["heading3"]="Heading3";
		Style["heading4"]="Heading4";
		Style["heading5"]="Heading5";
		Style["heading6"]="Heading6";
		Style["heading7"]="Heading7";
		Style["heading8"]="Heading8";
		Style["heading9"]="Heading9";
		Style["toc1"]="Toc1";
		Style["toc2"]="Toc2";
		Style["toc3"]="Toc3";
		Style["toc4"]="Toc4";
		Style["toc5"]="Toc5";
		Style["toc6"]="Toc6";
		Style["toc7"]="Toc7";
		Style["toc8"]="Toc8";
		Style["toc9"]="Toc9";
		Style["footnoteText"]="FootnoteText";
		Style["header"]="Header";
		Style["footer"]="Footer";
		Style["caption"]="Caption";
		Style["footnoteReference"]="FootnoteReference";
		Style["endnoteReference"]="EndnoteReference";
		Style["endnoteText"]="EndnoteText";
		Style["title"]="Title";
		Style["subtitle"]="Subtitle";
		Style["hyperlink"]="Hyperlink";
		Style["strong"]="Strong";
		Style["emphasis"]="Emphasis";
		Style["noSpacing"]="NoSpacing";
		Style["listParagraph"]="ListParagraph";
		Style["quote"]="Quote";
		Style["intenseQuote"]="IntenseQuote";
		Style["subtleEmphasis"]="SubtleEmphasis";
		Style["intenseEmphasis"]="IntenseEmphasis";
		Style["subtleReference"]="SubtleReference";
		Style["intenseReference"]="IntenseReference";
		Style["bookTitle"]="BookTitle";
		Style["bibliography"]="Bibliography";
		Style["tocHeading"]="TocHeading";
		Style["tableGrid"]="TableGrid";
		Style["plainTable1"]="PlainTable1";
		Style["plainTable2"]="PlainTable2";
		Style["plainTable3"]="PlainTable3";
		Style["plainTable4"]="PlainTable4";
		Style["plainTable5"]="PlainTable5";
		Style["tableGridLight"]="TableGridLight";
		Style["gridTable1Light"]="GridTable1Light";
		Style["gridTable1Light_Accent1"]="GridTable1Light_Accent1";
		Style["gridTable1Light_Accent2"]="GridTable1Light_Accent2";
		Style["gridTable1Light_Accent3"]="GridTable1Light_Accent3";
		Style["gridTable1Light_Accent4"]="GridTable1Light_Accent4";
		Style["gridTable1Light_Accent5"]="GridTable1Light_Accent5";
		Style["gridTable1Light_Accent6"]="GridTable1Light_Accent6";
		Style["gridTable2"]="GridTable2";
		Style["gridTable2_Accent1"]="GridTable2_Accent1";
		Style["gridTable2_Accent2"]="GridTable2_Accent2";
		Style["gridTable2_Accent3"]="GridTable2_Accent3";
		Style["gridTable2_Accent4"]="GridTable2_Accent4";
		Style["gridTable2_Accent5"]="GridTable2_Accent5";
		Style["gridTable2_Accent6"]="GridTable2_Accent6";
		Style["gridTable3"]="GridTable3";
		Style["gridTable3_Accent1"]="GridTable3_Accent1";
		Style["gridTable3_Accent2"]="GridTable3_Accent2";
		Style["gridTable3_Accent3"]="GridTable3_Accent3";
		Style["gridTable3_Accent4"]="GridTable3_Accent4";
		Style["gridTable3_Accent5"]="GridTable3_Accent5";
		Style["gridTable3_Accent6"]="GridTable3_Accent6";
		Style["gridTable4"]="GridTable4";
		Style["gridTable4_Accent1"]="GridTable4_Accent1";
		Style["gridTable4_Accent2"]="GridTable4_Accent2";
		Style["gridTable4_Accent3"]="GridTable4_Accent3";
		Style["gridTable4_Accent4"]="GridTable4_Accent4";
		Style["gridTable4_Accent5"]="GridTable4_Accent5";
		Style["gridTable4_Accent6"]="GridTable4_Accent6";
		Style["gridTable5Dark"]="GridTable5Dark";
		Style["gridTable5Dark_Accent1"]="GridTable5Dark_Accent1";
		Style["gridTable5Dark_Accent2"]="GridTable5Dark_Accent2";
		Style["gridTable5Dark_Accent3"]="GridTable5Dark_Accent3";
		Style["gridTable5Dark_Accent4"]="GridTable5Dark_Accent4";
		Style["gridTable5Dark_Accent5"]="GridTable5Dark_Accent5";
		Style["gridTable5Dark_Accent6"]="GridTable5Dark_Accent6";
		Style["gridTable6Colorful"]="GridTable6Colorful";
		Style["gridTable6Colorful_Accent1"]="GridTable6Colorful_Accent1";
		Style["gridTable6Colorful_Accent2"]="GridTable6Colorful_Accent2";
		Style["gridTable6Colorful_Accent3"]="GridTable6Colorful_Accent3";
		Style["gridTable6Colorful_Accent4"]="GridTable6Colorful_Accent4";
		Style["gridTable6Colorful_Accent5"]="GridTable6Colorful_Accent5";
		Style["gridTable6Colorful_Accent6"]="GridTable6Colorful_Accent6";
		Style["gridTable7Colorful"]="GridTable7Colorful";
		Style["gridTable7Colorful_Accent1"]="GridTable7Colorful_Accent1";
		Style["gridTable7Colorful_Accent2"]="GridTable7Colorful_Accent2";
		Style["gridTable7Colorful_Accent3"]="GridTable7Colorful_Accent3";
		Style["gridTable7Colorful_Accent4"]="GridTable7Colorful_Accent4";
		Style["gridTable7Colorful_Accent5"]="GridTable7Colorful_Accent5";
		Style["gridTable7Colorful_Accent6"]="GridTable7Colorful_Accent6";
		Style["listTable1Light"]="ListTable1Light";
		Style["listTable1Light_Accent1"]="ListTable1Light_Accent1";
		Style["listTable1Light_Accent2"]="ListTable1Light_Accent2";
		Style["listTable1Light_Accent3"]="ListTable1Light_Accent3";
		Style["listTable1Light_Accent4"]="ListTable1Light_Accent4";
		Style["listTable1Light_Accent5"]="ListTable1Light_Accent5";
		Style["listTable1Light_Accent6"]="ListTable1Light_Accent6";
		Style["listTable2"]="ListTable2";
		Style["listTable2_Accent1"]="ListTable2_Accent1";
		Style["listTable2_Accent2"]="ListTable2_Accent2";
		Style["listTable2_Accent3"]="ListTable2_Accent3";
		Style["listTable2_Accent4"]="ListTable2_Accent4";
		Style["listTable2_Accent5"]="ListTable2_Accent5";
		Style["listTable2_Accent6"]="ListTable2_Accent6";
		Style["listTable3"]="ListTable3";
		Style["listTable3_Accent1"]="ListTable3_Accent1";
		Style["listTable3_Accent2"]="ListTable3_Accent2";
		Style["listTable3_Accent3"]="ListTable3_Accent3";
		Style["listTable3_Accent4"]="ListTable3_Accent4";
		Style["listTable3_Accent5"]="ListTable3_Accent5";
		Style["listTable3_Accent6"]="ListTable3_Accent6";
		Style["listTable4"]="ListTable4";
		Style["listTable4_Accent1"]="ListTable4_Accent1";
		Style["listTable4_Accent2"]="ListTable4_Accent2";
		Style["listTable4_Accent3"]="ListTable4_Accent3";
		Style["listTable4_Accent4"]="ListTable4_Accent4";
		Style["listTable4_Accent5"]="ListTable4_Accent5";
		Style["listTable4_Accent6"]="ListTable4_Accent6";
		Style["listTable5Dark"]="ListTable5Dark";
		Style["listTable5Dark_Accent1"]="ListTable5Dark_Accent1";
		Style["listTable5Dark_Accent2"]="ListTable5Dark_Accent2";
		Style["listTable5Dark_Accent3"]="ListTable5Dark_Accent3";
		Style["listTable5Dark_Accent4"]="ListTable5Dark_Accent4";
		Style["listTable5Dark_Accent5"]="ListTable5Dark_Accent5";
		Style["listTable5Dark_Accent6"]="ListTable5Dark_Accent6";
		Style["listTable6Colorful"]="ListTable6Colorful";
		Style["listTable6Colorful_Accent1"]="ListTable6Colorful_Accent1";
		Style["listTable6Colorful_Accent2"]="ListTable6Colorful_Accent2";
		Style["listTable6Colorful_Accent3"]="ListTable6Colorful_Accent3";
		Style["listTable6Colorful_Accent4"]="ListTable6Colorful_Accent4";
		Style["listTable6Colorful_Accent5"]="ListTable6Colorful_Accent5";
		Style["listTable6Colorful_Accent6"]="ListTable6Colorful_Accent6";
		Style["listTable7Colorful"]="ListTable7Colorful";
		Style["listTable7Colorful_Accent1"]="ListTable7Colorful_Accent1";
		Style["listTable7Colorful_Accent2"]="ListTable7Colorful_Accent2";
		Style["listTable7Colorful_Accent3"]="ListTable7Colorful_Accent3";
		Style["listTable7Colorful_Accent4"]="ListTable7Colorful_Accent4";
		Style["listTable7Colorful_Accent5"]="ListTable7Colorful_Accent5";
		Style["listTable7Colorful_Accent6"]="ListTable7Colorful_Accent6";
	})(Style=Word.Style || (Word.Style={}));
	var DocumentPropertyType;
	(function (DocumentPropertyType) {
		DocumentPropertyType["string"]="String";
		DocumentPropertyType["number"]="Number";
		DocumentPropertyType["date"]="Date";
		DocumentPropertyType["boolean"]="Boolean";
	})(DocumentPropertyType=Word.DocumentPropertyType || (Word.DocumentPropertyType={}));
	var TapObjectType;
	(function (TapObjectType) {
		TapObjectType["chart"]="Chart";
		TapObjectType["smartArt"]="SmartArt";
		TapObjectType["table"]="Table";
		TapObjectType["image"]="Image";
		TapObjectType["slide"]="Slide";
		TapObjectType["ole"]="OLE";
		TapObjectType["text"]="Text";
	})(TapObjectType=Word.TapObjectType || (Word.TapObjectType={}));
	var FileContentFormat;
	(function (FileContentFormat) {
		FileContentFormat["base64"]="Base64";
		FileContentFormat["html"]="Html";
		FileContentFormat["ooxml"]="Ooxml";
	})(FileContentFormat=Word.FileContentFormat || (Word.FileContentFormat={}));
	var ErrorCodes;
	(function (ErrorCodes) {
		ErrorCodes["accessDenied"]="AccessDenied";
		ErrorCodes["generalException"]="GeneralException";
		ErrorCodes["invalidArgument"]="InvalidArgument";
		ErrorCodes["itemNotFound"]="ItemNotFound";
		ErrorCodes["notImplemented"]="NotImplemented";
	})(ErrorCodes=Word.ErrorCodes || (Word.ErrorCodes={}));
})(Word || (Word={}));
var Word;
(function (Word) {
	var RequestContext=(function (_super) {
		__extends(RequestContext, _super);
		function RequestContext(url) {
			var _this=_super.call(this, url) || this;
			_this.m_document=new Word.Document(_this, OfficeExtension.ObjectPathFactory.createGlobalObjectObjectPath(_this));
			_this._rootObject=_this.m_document;
			return _this;
		}
		Object.defineProperty(RequestContext.prototype, "document", {
			get: function () {
				return this.m_document;
			},
			enumerable: true,
			configurable: true
		});
		Object.defineProperty(RequestContext.prototype, "application", {
			get: function () {
				if (this.m_application==null) {
					this.m_application=new Word.Application(this, OfficeExtension.ObjectPathFactory.createNewObjectObjectPath(this, "Microsoft.WordServices.Application", false));
				}
				return this.m_application;
			},
			enumerable: true,
			configurable: true
		});
		return RequestContext;
	}(OfficeCore.RequestContext));
	Word.RequestContext=RequestContext;
	function run(arg1, arg2) {
		return OfficeExtension.ClientRequestContext._runBatch("Word.run", arguments, function () { return new Word.RequestContext(); });
	}
	Word.run=run;
})(Word || (Word={}));

OfficeExtension.Utility._doApiNotSupportedCheck=true;


