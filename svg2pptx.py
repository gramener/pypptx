"""
Converts an SVG shape into a Microsoft Office object, and saves as pptx file

"""
import re
from lxml import etree
from lxml.builder import ElementMaker
from pypptx import a, p, shape, color, nsmap

re_ns = re.compile(r'({.*?})?(.*)')

class Draw(object):
    def __init__(self, slide, width, height):
        self.slide = slide
        self.shapes = slide._element.find('.//p:spTree', namespaces=nsmap)
        # TODO: Replace 9999... with slide width and height
        self.x = lambda x: int(x * 9999999 / width)
        self.y = lambda y: int(y * 7777777 / height)

    def _shape_attrs(function):
        def wrapped(self, e):
            shape = function(self, e)
            keys = e.keys()
            if 'fill' in keys:
                shape.spPr.append(a.solidFill(color(srgbClr=e.get('fill')[1:])))
            # TODO: convert short color string to hex 
            if 'stroke' in keys:
                shape.spPr.append(a.ln(a.solidFill(color(srgbClr=e.get('stroke')[1:]))))
            # TODO: Add stroke width
            #if 'stroke-width' in keys:
            #    shape.spPr.append(a.ln(w='3175'))
            return shape
        return wrapped

    @_shape_attrs
    def circle(self, e):
        x = float(e.get('cx', 0))
        y = float(e.get('cy', 0))
        r = float(e.get('r', 0))
        shp = shape('ellipse', self.x(x - r), self.y(y - r),
                               self.x(2 * r), self.y(2 * r))
        self.shapes.append(shp)
        return shp

    @_shape_attrs
    def ellipse(self, e):
        x = float(e.get('cx', 0))
        y = float(e.get('cy', 0))
        rx = float(e.get('rx', 0))
        ry = float(e.get('ry', 0))
        shp = shape('ellipse', self.x(x - rx), self.y(y - ry),
                               self.x(2 * rx), self.y(2 * ry))
        self.shapes.append(shp)
        return shp

    @_shape_attrs
    def rect(self, e):
        shp = shape('rect',
            self.x(float(e.get('x', 0))),
            self.y(float(e.get('y', 0))),
            self.x(float(e.get('width', 0))),
            self.y(float(e.get('height', 0)))
        )
        self.shapes.append(shp)
        return shp

    @_shape_attrs
    def line(self, e):
        shp = shape('line',
            self.x(float(e.get('x1', 0))), self.y(float(e.get('y1', 0))),
            self.x(float(e.get('x2', 0))), self.y(float(e.get('y2', 0)))
        )
        self.shapes.append(shp)
        return shp

    def text(self, e):
        if not e.text:
            return
        # TODO: font-size (Autofit) and text orientation
        shp = shape('rect', self.x(float(e.get('x', 0))), self.y(float(e.get('y', 0))), self.x(0), self.y(0))
        shp.append(p.txBody(a.bodyPr(a.normAutofit(fontScale="62500", lnSpcReduction="20000"), anchor='ctr'),
            a.p(a.pPr(algn='ctr'),
                a.r(a.aPr(sz='1000'),
                    a.t(e.text)))))

        self.shapes.append(shp)
        return shp


def svg2mso(slide, svg, width=940, height=None):
    if width is not None and height is None:
        height = width * 3 / 4
    elif width is None and height is not None:
        width = height * 4 / 3

    # Convert tree into an lxml etree if it's not one
    if not hasattr(svg, 'iter'):
        svg = etree.parse(svg) if hasattr(svg, 'read') else etree.fromstring(svg)

    # Take all the tags and draw it
    draw = Draw(slide, width, height)
    valid_tags = set(tag for tag in dir(draw) if not tag.startswith('_'))
    for e in svg.iter(tag=etree.Element):
        match = re_ns.match(e.tag)
        if not match:
            continue

        tag = match.groups()[-1]
        if tag in valid_tags:
            getattr(draw, tag)(e)


if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description=__doc__.strip())
    parser.add_argument('svgfile')
    args = parser.parse_args()

    tree = etree.parse(open(args.svgfile))

    from pptx import Presentation

    Presentation = Presentation()
    blank_slidelayout = Presentation.slidelayouts[6]
    slide = Presentation.slides.add_slide(blank_slidelayout)

    svg2mso(slide, tree) 
    Presentation.save("test.pptx")