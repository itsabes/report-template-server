var ChartManager, DocUtils;

DocUtils = require('./docUtils');

module.exports = ChartManager = (function() {
  function ChartManager(zip, fileName) {
    this.zip = zip;
    this.fileName = fileName;
    this.endFileName = this.fileName.replace(/^.*?([a-z0-9]+)\.xml$/, "$1");
    this.relsLoaded = false;
  }


  /**
  	 * load relationships
  	 * @return {ChartManager} for chaining
   */

  ChartManager.prototype.loadChartRels = function() {

    /**
    		 * load file, save path
    		 * @param  {String} filePath path to current file
    		 * @return {Object}          file
     */
    var RidArray, content, file, loadFile, tag;
    loadFile = (function(_this) {
      return function(filePath) {
        _this.filePath = filePath;
        return _this.zip.files[_this.filePath];
      };
    })(this);
    file = loadFile("word/_rels/" + this.endFileName + ".xml.rels") || loadFile("word/_rels/document.xml.rels");
    if (file === void 0) {
      return;
    }
    content = DocUtils.decodeUtf8(file.asText());
    this.xmlDoc = DocUtils.Str2xml(content);
    RidArray = (function() {
      var _i, _len, _ref, _results;
      _ref = this.xmlDoc.getElementsByTagName('Relationship');
      _results = [];
      for (_i = 0, _len = _ref.length; _i < _len; _i++) {
        tag = _ref[_i];
        _results.push(parseInt(tag.getAttribute("Id").substr(3)));
      }
      return _results;
    }).call(this);
    this.maxRid = DocUtils.maxArray(RidArray);
    this.chartRels = [];
    this.relsLoaded = true;
    return this;
  };

  ChartManager.prototype.addChartRels = function(chartName) {
    if (!this.relsLoaded) {
      return;
    }
    this.maxRid++;
    this._addChartRelationship(this.maxRid, chartName);
    this._addChartContentType(chartName);
    this.zip.file(this.filePath, DocUtils.encodeUtf8(DocUtils.xml2Str(this.xmlDoc)), {});
    return this.maxRid;
  };


  /**
  	 * add relationship tag to relationships
  	 * @param {Number} id   relationship ID
  	 * @param {String} name target file name
   */

  ChartManager.prototype._addChartRelationship = function(id, name) {
    var newTag, relationships;
    relationships = this.xmlDoc.getElementsByTagName("Relationships")[0];
    newTag = this.xmlDoc.createElement('Relationship');
    newTag.namespaceURI = null;
    newTag.setAttribute('Id', "rId" + id);
    newTag.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart');
    newTag.setAttribute('Target', "charts/" + name + ".xml");
    return relationships.appendChild(newTag);
  };


  /**
  	 * add override to [Content_Types].xml
  	 * @param {String} name filename
   */

  ChartManager.prototype._addChartContentType = function(name) {
    var content, file, newTag, path, types, xmlDoc;
    path = '[Content_Types].xml';
    file = this.zip.files[path];
    content = DocUtils.decodeUtf8(file.asText());
    xmlDoc = DocUtils.Str2xml(content);
    types = xmlDoc.getElementsByTagName("Types")[0];
    newTag = xmlDoc.createElement('Override');
    newTag.namespaceURI = 'http://schemas.openxmlformats.org/package/2006/content-types';
    newTag.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml');
    newTag.setAttribute('PartName', "/word/charts/" + name + ".xml");
    types.appendChild(newTag);
    return this.zip.file(path, DocUtils.encodeUtf8(DocUtils.xml2Str(xmlDoc)), {});
  };

  return ChartManager;

})();
