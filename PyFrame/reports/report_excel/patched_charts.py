# coding=UTF-8
""" Copyright (c) 2010-2014 openpyxl
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.
#
# @license: http://www.opensource.org/licenses/mit-license.php
# @author: see AUTHORS file
"""

from openpyxl.xml.functions import (
    Element,
    SubElement,
    get_document_content
)
from openpyxl.xml.constants import (
    CHART_NS,
    DRAWING_NS,
    REL_NS,
    PKG_REL_NS
)
from openpyxl.compat import (
    iteritems,
    safe_string
)
from openpyxl.charts import (
    ErrorBar,
    GraphChart
)


class BaseChartWriter(object):
    """The baseChartWriter
    """

    series_type = '{%s}val' % CHART_NS

    def __init__(self, chart):
        """the metod init"""
        self.chart = chart
        self.root = Element("{%s}chartSpace" % CHART_NS)

    def write(self):
        """ write a chart """
        SubElement(self.root, '{%s}lang' % CHART_NS, {'val': self.chart.lang})
        self.write_chart()
        self._write_print_settings()
        self._write_shapes()

        return get_document_content(self.root)

    def write_chart(self):
        """write chart"""
        chart = SubElement(self.root, '{%s}chart' % CHART_NS)
        self._write_title(chart)
        self._write_layout(chart)
        self._write_legend(chart)
        SubElement(chart, '{%s}plotVisOnly' % CHART_NS, {'val': '1'})

    def _write_layout(self, element):
        """write layout"""
        chart = self.chart
        plot_area = SubElement(element, '{%s}plotArea' % CHART_NS)
        layout = SubElement(plot_area, '{%s}layout' % CHART_NS)
        mlayout = SubElement(layout, '{%s}manualLayout' % CHART_NS)
        SubElement(mlayout, '{%s}layoutTarget' % CHART_NS, {'val': 'inner'})
        SubElement(mlayout, '{%s}xMode' % CHART_NS, {'val': 'edge'})
        SubElement(mlayout, '{%s}yMode' % CHART_NS, {'val': 'edge'})
        SubElement(mlayout, '{%s}x' %
                   CHART_NS, {'val': safe_string(chart.margin_left)})
        SubElement(mlayout, '{%s}y' %
                   CHART_NS, {'val': safe_string(chart.margin_top)})
        SubElement(mlayout, '{%s}w' %
                   CHART_NS, {'val': safe_string(chart.width)})
        SubElement(mlayout, '{%s}h' %
                   CHART_NS, {'val': safe_string(chart.height)})

        chart_type = self.chart.TYPE
        subchart = SubElement(plot_area, '{%s}%s' % (CHART_NS, chart_type))
        self._write_options(subchart)
        self._write_series(subchart)
        if isinstance(chart, GraphChart):
            SubElement(subchart, '{%s}axId' %
                       CHART_NS, {'val': safe_string(chart.x_axis.id)})
            SubElement(subchart, '{%s}axId' %
                       CHART_NS, {'val': safe_string(chart.y_axis.id)})
            self._write_axis(
                plot_area, chart.x_axis, '{%s}%s' % (CHART_NS,
                                                     chart.x_axis.type))
            self._write_axis(
                plot_area, chart.y_axis, '{%s}%s' % (CHART_NS,
                                                     chart.y_axis.type))

    def _write_options(self, subchart):
        """write options"""
        pass

    def _write_title(self, chart):
        """write title"""
        if self.chart.title != '':
            title = SubElement(chart, '{%s}title' % CHART_NS)
            text = SubElement(title, '{%s}tx' % CHART_NS)
            rich = SubElement(text, '{%s}rich' % CHART_NS)
            SubElement(rich, '{%s}bodyPr' % DRAWING_NS)
            SubElement(rich, '{%s}lstStyle' % DRAWING_NS)
            parrafo = SubElement(rich, '{%s}p' % DRAWING_NS)
            p_pr = SubElement(parrafo, '{%s}pPr' % DRAWING_NS)
            SubElement(p_pr, '{%s}defRPr' % DRAWING_NS)
            rit = SubElement(parrafo, '{%s}r' % DRAWING_NS)
            SubElement(rit, '{%s}rPr' % DRAWING_NS, {'lang': self.chart.lang})
            SubElement(rit, '{%s}t' % DRAWING_NS).text = self.chart.title
            SubElement(title, '{%s}layout' % CHART_NS)

    def _write_axis_title(self, axis, a_x):
        """Write axis title"""

        if axis.title != '':
            title = SubElement(a_x, '{%s}title' % CHART_NS)
            text = SubElement(title, '{%s}tx' % CHART_NS)
            rich = SubElement(text, '{%s}rich' % CHART_NS)
            SubElement(rich, '{%s}bodyPr' % DRAWING_NS)
            SubElement(rich, '{%s}lstStyle' % DRAWING_NS)
            parrafo = SubElement(rich, '{%s}p' % DRAWING_NS)
            p_pr = SubElement(parrafo, '{%s}pPr' % DRAWING_NS)
            SubElement(p_pr, '{%s}defRPr' % DRAWING_NS)
            rit = SubElement(parrafo, '{%s}r' % DRAWING_NS)
            SubElement(rit, '{%s}rPr' % DRAWING_NS, {'lang': self.chart.lang})
            SubElement(rit, '{%s}t' % DRAWING_NS).text = axis.title
            SubElement(title, '{%s}layout' % CHART_NS)

    def _write_axis(self, plot_area, axis, label):
        """Write axis"""
        if self.chart.auto_axis:
            self.chart.compute_axes()

        a_x = SubElement(plot_area, label)
        SubElement(a_x, '{%s}axId' % CHART_NS, {'val': safe_string(axis.id)})

        scaling = SubElement(a_x, '{%s}scaling' % CHART_NS)
        SubElement(scaling, '{%s}orientation' %
                   CHART_NS, {'val': axis.orientation})
        if axis.delete_axis:
            SubElement(scaling, '{%s}' % CHART_NS, {'val': '1'})
        if axis.type == "valAx":
            SubElement(scaling, '{%s}max' %
                       CHART_NS, {'val': str(float(axis.max))})
            SubElement(scaling, '{%s}min' %
                       CHART_NS, {'val': str(float(axis.min))})

        SubElement(a_x, '{%s}axPos' % CHART_NS, {'val': axis.position})
        if axis.type == "valAx":
            SubElement(a_x, '{%s}majorGridlines' % CHART_NS)
            SubElement(
                a_x, '{%s}numFmt' % CHART_NS, {'formatCode': "General",
                                               'sourceLinked': '1'})
        self._write_axis_title(axis, a_x)
        SubElement(a_x, '{%s}tickLblPos' %
                   CHART_NS, {'val': axis.tick_label_position})
        SubElement(a_x, '{%s}crossAx' % CHART_NS, {'val': str(axis.cross)})
        SubElement(a_x, '{%s}crosses' % CHART_NS, {'val': axis.crosses})
        if axis.auto:
            SubElement(a_x, '{%s}auto' % CHART_NS, {'val': '1'})
        if axis.label_align:
            SubElement(a_x, '{%s}lblAlgn' % CHART_NS, {'val': axis.label_align})
        if axis.label_offset:
            SubElement(a_x, '{%s}lblOffset' %
                       CHART_NS, {'val': str(axis.label_offset)})
        if axis.type == "valAx":
            SubElement(a_x, '{%s}crossBetween' %
                       CHART_NS, {'val': axis.cross_between})
            SubElement(a_x, '{%s}majorUnit' %
                       CHART_NS, {'val': str(float(axis.unit))})

    def _write_series(self, subchart):
        """write series"""
        for i, serie in enumerate(self.chart):
            ser = SubElement(subchart, '{%s}ser' % CHART_NS)
            SubElement(ser, '{%s}idx' % CHART_NS, {'val': safe_string(i)})
            SubElement(ser, '{%s}order' % CHART_NS, {'val': safe_string(i)})

            if serie.title:
                t_x = SubElement(ser, '{%s}t_x' % CHART_NS)
                SubElement(t_x, '{%s}v' % CHART_NS).text = serie.title

            if serie.color:
                sppr = SubElement(ser, '{%s}spPr' % CHART_NS)
                _write_series_color_node(sppr, serie)
                d_pt = SubElement(ser, '{%s}dPt' % CHART_NS)
                SubElement(d_pt, '{%s}idx' % CHART_NS, {'val': '1'})
                SubElement(d_pt, '{%s}Bubble3D' % CHART_NS, {'val': '1'})
                sppr = SubElement(d_pt, '{%s}spPr' % CHART_NS)
                fillc = SubElement(sppr, '{%s}solidFill' % DRAWING_NS)
                SubElement(fillc, '{%s}srgbClr' % DRAWING_NS, {'val': 'FFFF00'})

            if serie.error_bar:
                _write_error_bar(ser, serie)

            if serie.labels:
                _write_series_labels(ser, serie)

            if serie.xvalues:
                _write_series_xvalues(ser, serie)

            val = SubElement(ser, self.series_type)
            _write_serial(val, serie.reference)


    def _write_legend(self, chart):
        """write legend"""
        if self.chart.show_legend:
            legend = SubElement(chart, '{%s}legend' % CHART_NS)
            SubElement(legend, '{%s}legendPos' %
                       CHART_NS, {'val': self.chart.legend.position})
            SubElement(legend, '{%s}layout' % CHART_NS)

    def _write_print_settings(self):
        """write print setting"""
        settings = SubElement(self.root, '{%s}printSettings' % CHART_NS)
        SubElement(settings, '{%s}headerFooter' % CHART_NS)
        margins = dict([(k, safe_string(v))
                        for (k, v) in iteritems(self.chart.print_margins)])
        SubElement(settings, '{%s}pageMargins' % CHART_NS, margins)
        SubElement(settings, '{%s}pageSetup' % CHART_NS)

    def _write_shapes(self):
        """write shapes"""
        if self.chart.shapes:
            SubElement(self.root, '{%s}userShapes' %
                       CHART_NS, {'{%s}id' % REL_NS: 'rId1'})


class PieChartWriter(BaseChartWriter):
    """The pie charts writer"""
    def _write_options(self, subchart):
        """write options"""
        pass

    def _write_axis(self, plot_area, axis, label):
        """Pie Charts have no axes, do nothing"""
        pass

    def _write_series(self, subchart):
        """write series"""
        for i, serie in enumerate(self.chart):
            ser = SubElement(subchart, '{%s}ser' % CHART_NS)
            SubElement(ser, '{%s}idx' % CHART_NS, {'val': safe_string(i)})
            SubElement(ser, '{%s}order' % CHART_NS, {'val': safe_string(i)})

            if serie.title:
                t_x = SubElement(ser, '{%s}tx' % CHART_NS)
                SubElement(t_x, '{%s}v' % CHART_NS).text = serie.title

            if serie.color:
                sppr = SubElement(ser, '{%s}spPr' % CHART_NS)
                _write_series_color_node(sppr, serie)

                d_pt = SubElement(ser, '{%s}dPt' % CHART_NS)
                SubElement(d_pt, '{%s}idx' % CHART_NS, {'val': '1'})
                sppr = SubElement(d_pt, '{%s}spPr' % CHART_NS)
                fillc = SubElement(sppr, '{%s}solidFill' % DRAWING_NS)
                SubElement(fillc, '{%s}srgbClr' % DRAWING_NS, {'val': 'FF0000'})

                d_pt = SubElement(ser, '{%s}dPt' % CHART_NS)
                SubElement(d_pt, '{%s}idx' % CHART_NS, {'val': '2'})
                sppr = SubElement(d_pt, '{%s}spPr' % CHART_NS)
                fillc = SubElement(sppr, '{%s}solidFill' % DRAWING_NS)
                SubElement(fillc, '{%s}srgbClr' % DRAWING_NS, {'val': '777777'})

                d_lbls = SubElement(ser, '{%s}dLbls' % CHART_NS)
                SubElement(d_lbls, '{%s}dLblPos' % DRAWING_NS,
                           {'val': 'outEnd'})
                SubElement(d_lbls, '{%s}showLegendKey' %
                           DRAWING_NS, {'val': '1'})
                SubElement(d_lbls, '{%s}showVal' % DRAWING_NS, {'val': '0'})
                SubElement(d_lbls, '{%s}showCatName' % DRAWING_NS, {'val': '0'})
                SubElement(d_lbls, '{%s}showSerName' % DRAWING_NS, {'val': '0'})
                SubElement(d_lbls, '{%s}showPercent' % DRAWING_NS, {'val': '1'})
                SubElement(d_lbls, '{%s}showBubbleSize' %
                           DRAWING_NS, {'val': '0'})
                SubElement(d_lbls, '{%s}showLeaderLines' %
                           DRAWING_NS, {'val': '0'})

            if serie.error_bar:
                _write_error_bar(ser, serie)

            if serie.labels:
                _write_series_labels(ser, serie)

            if serie.xvalues:
                _write_series_xvalues(ser, serie)

            val = SubElement(ser, self.series_type)
            _write_serial(val, serie.reference)


def _write_series_color_node(node, serie):
    """write series color
    """
    # edge color
    line = SubElement(node, '{%s}ln' % DRAWING_NS)
    fill = SubElement(line, '{%s}solidFill' % DRAWING_NS)
    SubElement(fill, '{%s}srgbClr' % DRAWING_NS, {'val': serie.color})

def _write_series_xvalues(node, serie):
    """write series xvalues"""
    raise NotImplemented("x values not possible for this chart type node:%s"
                         " serie:%s" %(node, serie))
def _write_serial(node, reference):
    """ write serial
    """
    is_ref = hasattr(reference, 'pos1')
    data_type = reference.data_type
    number_format = getattr(reference, 'number_format')

    mapping = {'n': {'ref': 'numRef', 'cache': 'numCache'},
               's': {'ref': 'strRef', 'cache': 'strCache'}}

    if is_ref:
        ref = SubElement(node, '{%s}%s' %
                         (CHART_NS, mapping[data_type]['ref']))
        SubElement(ref, '{%s}f' % CHART_NS).text = str(reference)
        data = SubElement(ref, '{%s}%s' %
                          (CHART_NS, mapping[data_type]['cache']))
        values = reference.values
    else:
        data = SubElement(node, '{%s}numLit' % CHART_NS)
        values = (1,)

    if data_type == 'n':
        SubElement(data, '{%s}formatCode' %
                   CHART_NS).text = number_format or 'General'

    SubElement(data, '{%s}ptCount' % CHART_NS, {'val': str(len(values))})
    for j, val in enumerate(values):
        point = SubElement(data, '{%s}pt' % CHART_NS, {'idx': str(j)})
        val = safe_string(val)
        SubElement(point, '{%s}v' % CHART_NS).text = val

def _write_series_labels(node, serie):
    """write series labels"""
    cat = SubElement(node, '{%s}cat' % CHART_NS)
    _write_serial(cat, serie.labels)

def _write_error_bar(node, serie):
    """write error bar"""
    flag = {ErrorBar.PLUS_MINUS: 'both',
            ErrorBar.PLUS: 'plus',
            ErrorBar.MINUS: 'minus'}

    e_b = SubElement(node, '{%s}errBars' % CHART_NS)
    SubElement(e_b, '{%s}errBarType' %
               CHART_NS, {'val': flag[serie.error_bar.type]})
    SubElement(e_b, '{%s}errValType' % CHART_NS, {'val': 'cust'})

    plus = SubElement(e_b, '{%s}plus' % CHART_NS)
    # apart from setting the type of data series the following has
    # no effect on the writer
    _write_serial(plus, serie.error_bar.reference)

    minus = SubElement(e_b, '{%s}minus' % CHART_NS)
    _write_serial(minus, serie.error_bar.reference)

def write_rels(drawing_id):
    """write rels"""
    root = Element("{%s}Relationships" % PKG_REL_NS)

    attrs = {'Id': 'rId1',
             'Type': '%s/chartUserShapes' % REL_NS,
             'Target': '../drawings/drawing%s.xml' % drawing_id}
    SubElement(root, '{%s}Relationship' % PKG_REL_NS, attrs)
    return get_document_content(root)

def _write_series_color(node):
    """write series color"""
    # fill color
    fillc = SubElement(node, '{%s}solidFill' % DRAWING_NS)
    SubElement(fillc, '{%s}srgbClr' % DRAWING_NS, {'val': '117711'})
