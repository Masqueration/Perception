from datetime import datetime


def base(name):
    return '''<!DOCTYPE html>
    <html>
    <head>
    <style>
    table {
        border-collapse: collapse;
        width: 100%;
    }
    
    th, td {
        text-align: left;
        padding: 8px;
    }
    
    tr:nth-child(even){background-color: #f2f2f2}
    
    th {
        background-color: #4286f4;
        color: white;
    }
    </style>
    </head>
    <body>
    <h2>Report (''' + name + ''') </h2> 
    <table>
      <tr>
        <th>Time</th>
        <th>From</th>
        <th>To</th>
        <th>Subject</th>'''


def end_html():
    return """</tr></table></body></html>"""


def add_cell(txt, link=False):
    if not isinstance(txt, str):
        try:
            txt = str(txt)
        except:
            txt = 'N/A'
    if link:
        txt = """<a href=""" + txt + """>""" + txt[-20:] + """</a>"""
    return """<td>""" + txt + """</td>"""
