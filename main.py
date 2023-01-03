import base64
import datetime
import html
import json
import os.path
import random
import re
import shutil
import sys
import tarfile
import urllib.error
import urllib.request
from io import StringIO
from pathlib import Path
from xml.etree import ElementTree

import docx
import docx2txt
import unidecode
import yaml
from docx.document import Document as doctwo
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
import xml.etree.ElementTree as ET

config = {
    "directories": {
        "input": "source",
        "output": "out",
        "themes": "themes",
    },
    "site": {
        "theme": "default",
        "domain": "wordflow.com"
    },
    "author": {
        "nickname": "ahr",
        "name": "author",
        "email": "you@me.com",
        "about": "I publish my word documents using wordflow!",
    },
    "generator": {
        "input": "docx",  # input language (md soon..)
    }
}
content = {}

styles = {
    "Title": "h1",
    "Heading 1": "h1",
    "Heading 2": "h2",
    "Heading 3": "h3",
    "Emphasis": "u",
    "Normal": "p",
    "List Paragraph": "li",
    "List Number": "li",
    "List Bullet": "li",
    "Intense Quote": "q"
}


def slugify(text):
    text = unidecode.unidecode(text).lower()
    r = re.sub(r'[\W_]+', '-', text)
    if r.endswith("-"):
        r = r[:len(r) - 1]
    return r


def htmltotext(htm):
    ret = html.unescape(htm)
    ret = ret.translate({
        8209: ord('-'),
        8220: ord('"'),
        8221: ord('"'),
        160: ord(' '),
    })
    ret = re.sub(r"\s", " ", ret, flags=re.MULTILINE)
    ret = re.sub("<br>|<br />|</p>|</div>|</h\d>", "\n", ret, flags=re.IGNORECASE)
    ret = re.sub('<.*?>', ' ', ret, flags=re.DOTALL)
    ret = re.sub(r"  +", " ", ret)
    ret = re.compile(r'<img.*?>').sub('', ret)
    return ret


def iter_block_items(parent):
    """
    Yield each paragraph and table child within *parent*, in document order.
    Each returned value is an instance of either Table or Paragraph. *parent*
    would most commonly be a reference to a main Document object, but
    also works for a _Cell object, which itself can contain paragraphs and tables.
    """
    if isinstance(parent, doctwo):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("something's not right")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def generatehtmltag(document):
    global styles
    run = document.add_run()
    font = run.font
    htmlstring = ""
    css = ""
    tag = styles[document.style.name]
    if document.paragraph_format.alignment is not None:
        css += "text-align: {0};".format(document.paragraph_format.alignment)
    if document.paragraph_format.left_indent is not None:
        css += "margin-left: {0};".format(document.paragraph_format.left_indent.pt * 0.1)
    if document.paragraph_format.right_indent is not None:
        css += "margin-right: {0};".format(document.paragraph_format.right_indent.pt * 0.1)
    if document.paragraph_format.line_spacing is not None:
        css += "line-height: {0};".format(document.paragraph_format.line_spacing)
    if font.size is not None:
        css += "font-size: {0};".format(font.size)
    if font.italic is not None and font.italic:
        css += "font-style: italic;"
    if font.bold is not None and font.bold:
        css += "font-weight: bold;"
    if font.underline is not None:
        if font.underline:
            css += "text-decoration-line: underline;"
        else:
            css += "text-decoration-line: " + font.underline + ";"
    if font.color.rgb is not None:
        css += "color: " + font.color.rgb + ";"
    htmlstring = "<" + tag + " style='" + css + "'>" + document.text + "</" + tag + ">"
    return "{0}".format(htmlstring)


def getcontent(file, document):
    global config
    global content
    html = ""
    if not os.path.exists(config["directories"]["output"]):
        os.mkdir(config["directories"]["output"])
    if not os.path.exists(config["directories"]["output"] + "/public"):
        os.mkdir(config["directories"]["output"] + "/public")
    if not os.path.exists(config["directories"]["output"] + "/public/images"):
        os.mkdir(config["directories"]["output"] + "/public/images")
    if config["generator"]["input"] == "docx":
        doc = docx.Document(file)
        doc_properties = doc.core_properties
        html = ""
        images = {}
        id = str(random.randint(10000, 99999))
        imagedir = "/public/images/" + slugify(document["file"]) + id
        if not os.path.exists(config["directories"]["output"] + imagedir):
            os.mkdir(config["directories"]["output"] + imagedir)
        docx2txt.process(file, config["directories"]["output"] + imagedir)
        for r in doc.part.rels.values():
            if isinstance(r._target, docx.parts.image.ImagePart):
                images[r.rId] = os.path.basename(r._target.partname)
        i = 0
        for block in iter_block_items(doc):
            if 'text' in str(block):
                for run in block.runs:
                    xmlstr = str(run.element.xml)
                if 'Graphic' in xmlstr:
                    for rId in images:
                        if rId in xmlstr:
                            html += "<img class='img-fluid' src='" + imagedir + "/" + images[rId] + "'>"
                if block.text is not None:
                    html += generatehtmltag(block)
            elif 'table' in str(block):
                tablehtml = "<table>"
                tab = doc.tables[i]
                for row in tab.rows:
                    tr = "<tr>"
                    for cell in row.cells:
                        tr += "<td>{0}</td>".format(cell.text)
                    tr += "</tr>"
                    tablehtml += tr
                tablehtml += "</table>"
                html += tablehtml
                i += 1
        document["id"] = id
        document["imagedir"] = imagedir
        if doc_properties.created == None:
            document["date"] = datetime.date.today().strftime("%B %d, %Y")
        else:
            document["date"] = doc_properties.created.strftime("%B %d, %Y")
        document["body"] = html
        content[document["file"]] = document


def scancontent():
    """
    Scan Posts
    Document() document.core_properties.created for date
    """
    global config
    global content
    if os.path.exists(config["directories"]["input"]):
        source = Path(config["directories"]["input"] + "/")
        files = source.glob("*")
        for file in files:
            if file.is_dir():
                doctype = file.name
                file = Path(config["directories"]["input"] + "/" + file.name).glob("*." + config["generator"]["input"])
                for filecontent in file:
                    document = {
                        "type": doctype,
                        "file": filecontent.name.split(".")[0],
                        "title": filecontent.name.split(".")[0],
                        "body": "",
                    }
                    if not document["file"] in content:
                        document.update(config["author"])
                        document.update(config["site"])
                        getcontent(filecontent, document)
            else:
                print("Found misplaced file " + file.name + " please categorize your documents correctly. Skipping.")
    json_object = json.dumps(content, indent=4)
    with open('generated_output.json', 'w') as file:
        file.write(json_object)


def parsetemplate(input, type):
    """
    :input: string whom will be added into parsed template
    :type: page, post, category, search, profile
    :rtype: string
    """
    global config
    themefile = config['directories']['themes'] + "/" + config['site']['theme'] + "/" + type + ".html"
    if os.path.exists(themefile):
        p = re.compile('(\[\[([a-z]+)\]\])')
        output = str(open(themefile).read())
        matches = p.findall(output)
        for placeholder, token in matches:
            if token in input:
                output = output.replace(placeholder, str(input[token]))
        return output
    else:
        print("Warning!!! Template not found...")


def generatehomepage():
    global config
    global content
    homecontent = {}
    homecontent.update(config["author"])
    homecontent.update(config["site"])
    homecontent["body"] = ""
    for post in content:
        if content[post]["type"] == "post":
            body = (content[post]["body"][:75] + '..') if len(content[post]["body"]) > 75 else content[post]["body"]
            homecontent["body"] += '<div class="card post-item bg-transparent border-0 mb-5">'
            homecontent["body"] += '<div class="card-body px-0">'
            homecontent["body"] += '<h2 class="card-title">'
            homecontent["body"] += '<a class="text-white opacity-75-onHover" href="/post/{0}">{1}</a>'.format(
                slugify(content[post]["file"]), content[post]["title"])
            homecontent["body"] += '</h2>'
            homecontent["body"] += '<ul class="post-meta mt-3">'
            homecontent["body"] += '<li class="d-inline-block mr-3">'
            homecontent["body"] += '<span class="fas fa-clock text-primary"></span>'
            homecontent["body"] += '<a class="ml-1" href="#">{0}</a>'.format(content[post]["date"])
            homecontent["body"] += '</li>'
            homecontent["body"] += '<li class="d-inline-block">'
            homecontent["body"] += '<span class="fas fa-list-alt text-primary"></span>'
            homecontent["body"] += '<a class="ml-1" href="#">{0}</a>'.format(config["author"]["name"])
            homecontent["body"] += '</li>'
            homecontent["body"] += '</ul>'
            homecontent["body"] += '<p class="card-text my-4">{0}</p>'.format(htmltotext(body))
            homecontent["body"] += '<a href="/post/{0}.html" class="btn btn-primary">Read More</a>'.format(
                slugify(content[post]["file"]))
            homecontent["body"] += '</div>'
            homecontent["body"] += '</div>'
    filename = config["directories"]["output"] + "/index.html"
    outfile = open(filename, "w")
    outfile.write(parsetemplate(homecontent, "home"))
    outfile.close()


def generatehtml():
    scancontent()
    generatehomepage()
    for doc in content:
        document = content[doc]
        if not os.path.exists(config["directories"]["output"] + "/" + document["type"]):
            os.mkdir(config["directories"]["output"] + "/" + document["type"])
        filename = config["directories"]["output"] + "/" + document["type"] + "/" + slugify(document["file"]) + ".html"
        outfile = open(filename, "w")
        outfile.write(parsetemplate(document, document["type"]))
        outfile.close()
    print("Checking assets for the theme...")
    if os.path.exists(config["directories"]["themes"] + "/" + config["site"]["theme"] + "/assets"):
        if os.path.exists(config["directories"]["output"] + "/public/assets"):
            shutil.rmtree(config["directories"]["output"] + "/public/assets")
        shutil.copytree(config["directories"]["themes"] + "/" + config["site"]["theme"] + "/assets",
                        config["directories"]["output"] + "/public/assets")


def downloadtheme(name):
    global config
    url = "https://api.github.com/repos/devsimsek/WordFlow_themes/tarball/" + name + "_theme"
    try:
        status = urllib.request.urlopen(url)
    except urllib.error.HTTPError:
        print("Theme " + name + " not found.")
        return
    if not os.path.exists(config["directories"]["themes"] + "/" + name):
        if not os.path.exists("temp"):
            os.mkdir("temp")
        if not os.path.exists(config["directories"]["themes"] + "/" + name):
            os.mkdir(config["directories"]["themes"] + "/" + name)
        urllib.request.urlretrieve(url, "temp/" + name + "_theme.tar.gz")
        theme = tarfile.open("temp/" + name + "_theme.tar.gz")
        theme.extractall(config["directories"]["themes"] + "/" + name)
        extractedfile = os.path.commonprefix(theme.getnames())
        theme.close()
        for file in Path(config["directories"]["themes"] + "/" + name + "/" + extractedfile).glob("*"):
            shutil.move(file, config["directories"]["themes"] + "/" + name)
        shutil.rmtree(config["directories"]["themes"] + "/" + name + "/" + extractedfile)
        shutil.rmtree("temp")
        print("Theme installed.")
    else:
        print("Selected theme already exists. Want to reinstall?")
        val = input("(yes, no)> ")
        if val != "yes":
            print("Skipping theme installation.")
        else:
            print("Reinstalling...")
            shutil.rmtree(config["directories"]["themes"] + "/" + name)
            downloadtheme(name)


def clearinstallation():
    global config
    for directory in config["directories"]:
        print("removing " + config["directories"][directory])
        shutil.rmtree(config["directories"][directory])
    if os.path.exists("config.yaml"):
        os.remove("config.yaml")
    if os.path.exists("generated_output.json"):
        os.remove("generated_output.json")


def clearcontent():
    global config
    for directory in config["directories"]:
        if directory == "input":
            continue
        if directory == "themes":
            continue
        print("removing " + config["directories"][directory])
        shutil.rmtree(config["directories"][directory])
        if not os.path.exists(config["directories"][directory]):
            os.mkdir(config["directories"][directory])
        os.remove("generated_output.json")


def initapp():
    """
    Initialize wordflow application
    """
    global config
    print("Welcome to the WordFlow initializer.")
    print("Checking configuration")
    if not os.path.exists("config.yaml"):
        for key in config:
            i = 1
            for opt in config[key]:
                print(
                    "--- " + key.capitalize() + " Configuration (" + str(i) + " of " + str(len(config[key])) + ") ---")
                print("Configuring " + opt + " field.")
                val = input("Value (default: " + config[key][opt] + ")> ")
                if not val == "":
                    config[key][opt] = val
                i += 1
        print("Configuration completed.")
        if not os.path.exists("config.yaml"):
            with open("config.yaml", "w") as file:
                try:
                    yaml.dump(config, file)
                    print("Configuration saved. You can create your documents now.")
                except yaml.YAMLError as exception:
                    print(exception)
        else:
            print("Operation failed. Configuration already exists!")
            val = input(
                "Want to clean install WordFlow? (this will remove every configuration and files.) (yes or no)> ")
            if val != "yes":
                print("Bye :)")
            else:
                clearinstallation()
    else:
        print("Configuration found. Skipping.")
    print("Checking directories")
    for directory in config["directories"]:
        if not os.path.exists(config["directories"][directory]):
            print(config["directories"][directory] + " not exists. Creating.")
            os.mkdir(config["directories"][directory])
        else:
            print(directory + " exists.")
    if os.path.exists(config["directories"]["input"]):
        if not os.path.exists(config["directories"]["input"] + "/post"):
            os.mkdir(config["directories"]["input"] + "/post")
        if not os.path.exists(config["directories"]["input"] + "/page"):
            os.mkdir(config["directories"]["input"] + "/page")
    if config["site"]["theme"] == "default":
        print("Installing Theme")
        downloadtheme(config["site"]["theme"])

    print("Application should be initialized correctly. Thanks for using WordFlow.")
    exit(1)


def argvparser():
    args = sys.argv[1:]
    for arg in args:
        if arg == "init" or arg == "-init":
            initapp()
        elif arg == "generate" or arg == "gen":
            generatehtml()
        elif arg == "clear":
            val = input(
                "Want to clean install WordFlow? (this will remove every configuration and files.) (yes or no)> ")
            if val != "yes":
                print("Bye :)")
            else:
                clearinstallation()
        elif arg == "installtheme" or arg == "theme":
            name = input("Theme name> ")
            downloadtheme(name)
        elif arg == "scan":
            scancontent()
        elif arg == "clearcontent" or arg == "-cc":
            val = input(
                "Want to wipe all generated content? (yes or no)> ")
            if val != "yes":
                print("Bye :)")
            else:
                clearcontent()


def wordflow():
    if not os.path.exists("config.yaml"):
        print("Warning: Configuration file not found. Launching initializer.")
        initapp()
    else:
        global config
        with open("config.yaml") as file:
            try:
                config = yaml.safe_load(file)
            except yaml.YAMLError as exception:
                print(exception)


if __name__ == "__main__":
    wordflow()
    argvparser()
else:
    print("Illegal Launch Option")
