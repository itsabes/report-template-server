var ChartMaker, ChartManager, ChartModule, SubContent, fs;

SubContent = require('docxtemplater').SubContent;

ChartManager = require('./chartManager');

ChartMaker = require('./chartMaker');

fs = require('fs');

ChartModule = (function() {

  /**
  	 * self name for self-identification, variable for fast changing;
  	 * @type {String}
   */
  ChartModule.prototype.name = 'chart';


  /**
  	 * initialize options with empty object if not recived
  	 * @manager = ModuleManager instance
  	 * @param  {Object} @options params for the module
   */

  function ChartModule(options) {
    this.options = options != null ? options : {};
  }

  ChartModule.prototype.handleEvent = function(event, eventData) {
    var gen;
    if (event === 'rendering-file') {
      this.renderingFileName = eventData;
      gen = this.manager.getInstance('gen');
      this.chartManager = new ChartManager(gen.zip, this.renderingFileName);
      return this.chartManager.loadChartRels();
    } else if (event === 'rendered') {
      return this.finished();
    }
  };

  ChartModule.prototype.get = function(data) {
    var templaterState;
    if (data === 'loopType') {
      templaterState = this.manager.getInstance('templaterState');
      if (templaterState.textInsideTag[0] === '$') {
        return this.name;
      }
    }
    return null;
  };

  ChartModule.prototype.handle = function(type, data) {
    if (type === 'replaceTag' && data === this.name) {
      this.replaceTag();
    }
    return null;
  };

  ChartModule.prototype.finished = function() {};

  ChartModule.prototype.on = function(event, data) {
    if (event === 'error') {
      throw data;
    }
  };

  ChartModule.prototype.replaceBy = function(text, outsideElement) {
    var subContent, templaterState, xmlTemplater;
    xmlTemplater = this.manager.getInstance('xmlTemplater');
    templaterState = this.manager.getInstance('templaterState');
    subContent = new SubContent(xmlTemplater.content).getInnerTag(templaterState).getOuterXml(outsideElement);
    return xmlTemplater.replaceXml(subContent, text);
  };

  ChartModule.prototype.convertPixelsToEmus = function(pixel) {
    return Math.round(pixel * 9525);
  };

  ChartModule.prototype.extendDefaults = function(options) {
    var deepMerge, defaultOptions, result;
    deepMerge = function(target, source) {
      var key, next, original;
      for (key in source) {
        original = target[key];
        next = source[key];
        if (original && next && typeof next === 'object') {
          deepMerge(original, next);
        } else {
          target[key] = next;
        }
      }
      return target;
    };
    defaultOptions = {
      width: 5486400 / 9525,
      height: 3200400 / 9525,
      grid: true,
      border: true,
      title: false,
      legend: {
        position: 'r'
      },
      axis: {
        x: {
          orientation: 'minMax',
          min: void 0,
          max: void 0,
          type: void 0,
          date: {
            format: 'unix',
            code: 'mm/yy',
            unit: 'months',
            step: '1'
          }
        },
        y: {
          orientation: 'minMax',
          mix: void 0,
          max: void 0
        }
      }
    };
    result = deepMerge({}, defaultOptions);
    result = deepMerge(result, options);
    return result;
  };

  ChartModule.prototype.convertUnixTo1900 = function(chartData, axName) {
    var convertOption, data, line, unixTo1900, _i, _j, _len, _len1, _ref, _ref1;
    unixTo1900 = function(value) {
      return Math.round(value / 86400 + 25569);
    };
    convertOption = function(name) {
      if (chartData.options.axis[axName][name]) {
        return chartData.options.axis[axName][name] = unixTo1900(chartData.options.axis[axName][name]);
      }
    };
    convertOption('min');
    convertOption('max');
    _ref = chartData.lines;
    for (_i = 0, _len = _ref.length; _i < _len; _i++) {
      line = _ref[_i];
      _ref1 = line.data;
      for (_j = 0, _len1 = _ref1.length; _j < _len1; _j++) {
        data = _ref1[_j];
        data[axName] = unixTo1900(data[axName]);
      }
    }
    return chartData;
  };

  ChartModule.prototype.replaceTag = function() {
    var ax, chart, chartData, chartId, filename, gen, imageRels, name, newText, options, scopeManager, tag, tagXml, templaterState;
    scopeManager = this.manager.getInstance('scopeManager');
    templaterState = this.manager.getInstance('templaterState');
    gen = this.manager.getInstance('gen');
    tag = templaterState.textInsideTag.substr(1);
    chartData = scopeManager.getValueFromScope(tag);
    if (chartData == null) {
      return;
    }
    filename = tag + (this.chartManager.maxRid + 1);
    imageRels = this.chartManager.loadChartRels();
    if (!imageRels) {
      return;
    }
    chartId = this.chartManager.addChartRels(filename);
    options = this.extendDefaults(chartData.options);
    for (name in options.axis) {
      ax = options.axis[name];
      if (ax.type === 'date' && ax[ax.type].format === 'unix') {
        chartData = this.convertUnixTo1900(chartData, name);
      }
    }
    chart = new ChartMaker(gen.zip, options);
    chart.makeChartFile(chartData.lines);
    chart.writeFile(filename);
    tagXml = this.manager.getInstance('xmlTemplater').tagXml;
    newText = this.getChartXml({
      chartID: chartId,
      width: this.convertPixelsToEmus(options.width),
      height: this.convertPixelsToEmus(options.height)
    });
    return this.replaceBy(newText, tagXml);
  };

  ChartModule.prototype.getChartXml = function(_arg) {
    var chartID, height, width;
    chartID = _arg.chartID, width = _arg.width, height = _arg.height;
    return "<w:drawing>\n	<wp:inline distB=\"0\" distL=\"0\" distR=\"0\" distT=\"0\">\n		<wp:extent cx=\"" + width + "\" cy=\"" + height + "\"/>\n		<wp:effectExtent b=\"0\" l=\"0\" r=\"0\" t=\"0\"/>\n		<wp:docPr id=\"1\" name=\"Диаграмма 1\"/>\n		<wp:cNvGraphicFramePr/>\n		<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">\n			<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">\n				<c:chart r:id=\"rId" + chartID + "\" xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>\n			</a:graphicData>\n		</a:graphic>\n	</wp:inline>\n</w:drawing>";
  };

  return ChartModule;

})();

module.exports = ChartModule;
