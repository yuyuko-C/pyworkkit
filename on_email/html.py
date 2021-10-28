
class Html_Maker:
    @classmethod
    def read_text_file(cls, path: str):
        ret = ''
        with open(path, 'r') as f:
            while True:
                line = f.readline()
                if not line:
                    break
                line = line.replace(
                    '\t', '<span style="white-space:pre">\t</span>')
                line = '<br>' if line == '\n' else line
                ret += '<div>'+line+'</div>\n'
        return ret

    @classmethod
    def hyper_link(cls, href: str, display: str):
        return '<a href="{}">'.format(href) + display + '</a>'

    @classmethod
    def image(cls, scr: str, width: int = None, height: int = None):
        info = []
        info.append('src="{}"'.format(scr))
        if width:
            width = 'width="{}px"'.format(width)
            info.append(width)
        if height:
            height = 'height="{}px"'.format(height)
            info.append(height)

        return '<img '+' '.join(info) + ' />'

    @classmethod
    def paragraph(cls, text: str):
        if text:
            return '<div>'+text+'</div>'
        else:
            return '<br>'
