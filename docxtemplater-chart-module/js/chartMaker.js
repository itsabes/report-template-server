var ChartMaker, DocUtils;

DocUtils = require("./docUtils");

module.exports = ChartMaker = (function() {
  ChartMaker.prototype.getTemplateTop = function() {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n	<c:lang val="ru-RU"/>\n	<c:chart>\n		' +
      (this.options.title ? "" : '<c:autoTitleDeleted val="1"/>') +
      '\n		<c:plotArea>\n			<c:layout/>\n			<c:lineChart>\n				<c:grouping val="standard"/>'
    );
  };

  ChartMaker.prototype.getFormatCode = function() {
    if (this.options.axis.x.type === "date") {
      return "<c:formatCode>m/d/yyyy</c:formatCode>";
    } else {
      return "";
    }
  };

  ChartMaker.prototype.getLineTemplate = function(line, i, count) {
    var shapeEnum = ["diamond", "square", "circle", "triangle"];
    var shapeSize = [10, 4, 4, 4];
    var elem, result, _i, _j, _len, _len1, _ref, _ref1;
    result =
      '<c:ser><c:marker><c:symbol val="' + shapeEnum[i%4] + '"/><c:size val="' + shapeSize[i%4] + '"/></c:marker>\n	<c:idx val="' +
      i +
      '"/>\n	<c:order val="' +
      i +
      '"/>\n	<c:tx>\n		<c:v>' +
      line.name +
      "</c:v>\n	</c:tx>\n	<c:cat>\n\n		<c:" +
      this.ref +
      ">\n			<c:" +
      this.cache +
      ">\n				" +
      this.getFormatCode() +
      '\n				<c:ptCount val="' +
      line.data.length +
      '"/>\n	';
    _ref = line.data;
    for (i = _i = 0, _len = _ref.length; _i < _len; i = ++_i) {
      elem = _ref[i];
      result += '<c:pt idx="' + i + '">\n	<c:v>' + elem.x + "</c:v>\n</c:pt>";
    }
    result +=
      "		</c:" +
      this.cache +
      ">\n	</c:" +
      this.ref +
      '>\n</c:cat>\n<c:val>\n	<c:numRef>\n		<c:numCache>\n			<c:formatCode>General</c:formatCode>\n			<c:ptCount val="' +
      line.data.length +
      '"/>';
    _ref1 = line.data;
    for (i = _j = 0, _len1 = _ref1.length; _j < _len1; i = ++_j) {
      elem = _ref1[i];
      result += '<c:pt idx="' + i + '">\n	<c:v>' + elem.y + "</c:v>\n</c:pt>";
    }
    result += "			</c:numCache>\n		</c:numRef>\n	</c:val>\n</c:ser>";
    return result;
  };

  ChartMaker.prototype.id1 = 142309248;

  ChartMaker.prototype.id2 = 142310784;

  ChartMaker.prototype.getScaling = function(opts) {
    return (
      '<c:scaling>\n	<c:orientation val="' +
      opts.orientation +
      '"/>\n	' +
      (opts.max !== void 0 ? '<c:max val="' + opts.max + '"/>' : "") +
      "\n	" +
      (opts.min !== void 0 ? '<c:min val="' + opts.min + '"/>' : "") +
      "\n</c:scaling>"
    );
  };

  ChartMaker.prototype.getAxOpts = function() {
    return (
      '<c:axId val="' +
      this.id1 +
      '"/>\n' +
      this.getScaling(this.options.axis.x) +
      '\n<c:axPos val="b"/>\n<c:tickLblPos val="nextTo"/>\n<c:txPr>\n	<a:bodyPr/>\n	<a:lstStyle/>\n	<a:p>\n		<a:pPr>\n			<a:defRPr sz="800"/>\n		</a:pPr>\n		<a:endParaRPr lang="ru-RU"/>\n	</a:p>\n</c:txPr>\n<c:crossAx val="' +
      this.id2 +
      '"/>\n<c:crosses val="autoZero"/>\n<c:auto val="1"/>\n<c:lblOffset val="100"/>'
    );
  };

  ChartMaker.prototype.getCatAx = function() {
    return (
      "<c:catAx>\n	" + this.getAxOpts() + '\n	<c:lblAlgn val="ctr"/>\n</c:catAx>'
    );
  };

  ChartMaker.prototype.getDateAx = function() {
    return (
      "<c:dateAx>\n	" +
      this.getAxOpts() +
      '\n	<c:delete val="0"/>\n	<c:numFmt formatCode="' +
      this.options.axis.x.date.code +
      '" sourceLinked="0"/>\n	<c:majorTickMark val="out"/>\n	<c:minorTickMark val="none"/>\n	<c:baseTimeUnit val="days"/>\n	<c:majorUnit val="' +
      this.options.axis.x.date.step +
      '"/>\n	<c:majorTimeUnit val="' +
      this.options.axis.x.date.unit +
      '"/>\n</c:dateAx>'
    );
  };

  ChartMaker.prototype.getBorder = function() {
    if (!this.options.border) {
      return "<c:spPr>\n	<a:ln>\n		<a:noFill/>\n	</a:ln>\n</c:spPr>";
    } else {
      return "";
    }
  };

  ChartMaker.prototype.getTemplateBottom = function() {
    var result;
    result =
      '	<c:marker val="1"/>\n	<c:axId val="' +
      this.id1 +
      '"/>\n	<c:axId val="' +
      this.id2 +
      '"/>\n</c:lineChart>';
    switch (this.options.axis.x.type) {
      case "date":
        result += this.getDateAx();
        break;
      default:
        result += this.getCatAx();
    }
    result +=
      '			<c:valAx>\n				<c:axId val="' +
      this.id2 +
      '"/>\n				' +
      this.getScaling(this.options.axis.y) +
      '\n				<c:axPos val="l"/>\n				' +
      (this.options.grid ? "<c:minorGridlines/>" : "") +
      (this.options.xTitle
        ? '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>' +
          this.options.xTitle +
          "</a:t></a:r></a:p></c:rich></c:tx></c:title>"
        : "") +
      '\n				<c:numFmt formatCode="General" sourceLinked="1"/>\n				<c:tickLblPos val="nextTo"/>\n				<c:txPr>\n					<a:bodyPr/>\n					<a:lstStyle/>\n					<a:p>\n						<a:pPr>\n							<a:defRPr sz="600"/>\n						</a:pPr>\n						<a:endParaRPr lang="ru-RU"/>\n					</a:p>\n				</c:txPr>\n				<c:crossAx val="' +
      this.id1 +
      '"/>\n				<c:crosses val="autoZero"/>\n				<c:crossBetween val="between"/>\n			</c:valAx>\n		<c:dTable><c:showHorzBorder val="1"/><c:showVertBorder val="1"/><c:showOutline val="1"/><c:showKeys val="1"/></c:dTable> </c:plotArea>\n		<c:plotVisOnly val="1"/>\n	</c:chart>\n	' +
      this.getBorder() +
      "\n</c:chartSpace>";
    return result;
  };

  function ChartMaker(zip, options) {
    this.zip = zip;
    this.options = options;
    if (this.options.axis.x.type === "date") {
      this.ref = "numRef";
      this.cache = "numCache";
    } else {
      this.ref = "strRef";
      this.cache = "strCache";
    }
  }

  ChartMaker.prototype.makeChartFile = function(lines) {
    var i, line, result, _i, _len;
    result = this.getTemplateTop();
    for (i = _i = 0, _len = lines.length; _i < _len; i = ++_i) {
      line = lines[i];
      result += this.getLineTemplate(line, i, _len);
    }
    result += this.getTemplateBottom();
    this.chartContent = result;
    return this.chartContent;
  };

  ChartMaker.prototype.writeFile = function(path) {
    this.zip.file("word/charts/" + path + ".xml", this.chartContent, {});
  };

  return ChartMaker;
})();
