"""
Converts an SVG shape into a Microsoft Office object, and saves as pptx file

"""
import re
from lxml import etree
from lxml.builder import ElementMaker
from pypptx import a, p, shape, color, nsmap, cust_shape
from color import rgba

re_ns = re.compile(r'({.*?})?(.*)')
re_path = re.compile(r'[mMzZlLhHvVcCsSqQtTaA]|[\+\-]?[\d\.e]+')

def msclr(color):
    r, g, b, a = rgba(color)
    return '%02x%02x%02x' % (255*r, 255*g, 255*b)


class Draw(object):
    def __init__(self, slide, width, height):
        self.slide = slide
        self.shapes = slide._element.find('.//p:spTree', namespaces=nsmap)
        # TODO: Replace 9999... with slide width and height
        self.x = lambda x: int(float(x) * 9999999 / width)
        self.y = lambda y: int(float(y) * 7777777 / height)

    def _shape_attrs(function):
        def wrapped(self, e):
            shape = function(self, e)
            keys = e.keys()
            if 'fill' in keys:
                shape.spPr.append(a.solidFill(color(srgbClr=msclr(e.get('fill')))))
            if 'stroke' in keys:
                shape.spPr.append(a.ln(a.solidFill(color(srgbClr=msclr(e.get('stroke'))))))
            if not 'stroke' in keys:
                shape.spPr.append(a.ln(a.solidFill(color(srgbClr='000000'))))
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
            self.x(e.get('x', 0)),
            self.y(e.get('y', 0)),
            self.x(e.get('width', 0)),
            self.y(e.get('height', 0))
        )
        self.shapes.append(shp)
        return shp

    @_shape_attrs
    def line(self, e):
        x1, y1 = self.x(e.get('x1', 0)), self.y(e.get('y1', 0))
        x2, y2 = self.x(e.get('x2', 0)), self.y(e.get('y2', 0))
        shp = shape('line', x1, y1, x2 - x1, y2 - y1)
        self.shapes.append(shp)
        return shp

    def text(self, e):
        keys = e.keys()
        if not e.text:
            return
        shp = shape('rect', self.x(e.get('x', 0)), self.y(e.get('y', 0)), self.x(0), self.y(0))
        if 'transform' in keys:
            shp.append(p.txBody(a.bodyPr(a.normAutofit(fontScale="62500", lnSpcReduction="20000"),
                a.scene3d(a.camera(a.rot(lat='0', lon='0', rev=str(int(e.get('transform')[8:11])*60000)),
                    prst='orthographicFront'), a.lightRig(rig='threePt', dir='t')),
                anchor='ctr', wrap='none'),
            a.p(a.pPr(algn='r'),
                a.r(a.t(e.text)))))
        else:
            shp.append(p.txBody(a.bodyPr(a.normAutofit(fontScale="62500", lnSpcReduction="20000"),
                anchor='ctr', wrap='none'),
            a.p(a.pPr(algn='ctr'),
                a.r(a.t(e.text)))))

        self.shapes.append(shp)
        return shp

    @_shape_attrs
    def path(self, e):
        pathstr = re_path.findall(e.get('d', ''))
        n, length, cmd, relative, shp = 0, len(pathstr), None, False, None
        x1, y1 = 0, 0
        xy = lambda n: (float(pathstr[n]) + (x1 if relative else 0),
                        float(pathstr[n + 1]) + (y1 if relative else 0))

        shp = cust_shape(x1, y1, 100000, 100000)
        path = a.path(w="100000", h="100000")
        shp.find('.//a:custGeom', namespaces=nsmap).append(
            a.pathLst(path))

        while n < length:
            if pathstr[n].lower() in 'mzlhvcsqta':
                cmd = pathstr[n].lower()
                relative = str.islower(pathstr[n])
                n += 1

            if cmd == 'm':
                x1, y1 = xy(n)
                path.append(a.moveTo(a.pt(x=str(self.x(x1)), y=str(self.y(y1)))))
                n += 2

            elif cmd == 'l':
                x1, y1 = xy(n)
                path.append(a.lnTo(a.pt(x=str(self.x(x1)), y=str(self.y(y1)))))
                n += 2

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
