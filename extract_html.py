import os
import zipfile
import tinycss
from bs4 import BeautifulSoup

def extract_html(file_name):
    """Extracts html as a string from the google doc zip file"""
    zf = zipfile.ZipFile(file_name)
    files = zf.namelist()
    html_file_name = None
    for f in files:
        if f.upper().endswith(".HTML"):
            html_file_name = f
    if html_file_name is not None:
        html_doc = zf.read(html_file_name)
        return html_doc
    else:
        return None
    
def extract_pngs(file_name, output_dir):
    """Extracts any embedded images from the zip file to output_dir"""
    zf = zipfile.ZipFile(file_name)
    files = zf.namelist()
    for f in files:
        if f.upper().endswith(".PNG"):
            output_name = os.path.join(output_dir, f)
            img_dir = os.path.split(output_name)[0]
            if not os.path.exists(img_dir):
                os.makedirs(img_dir)
            data = zf.read(f)
            f_out = open(output_name, "wb")
            f_out.write(data)
            f_out.close()
            

def get_html_soup(file_name):
    """Creates BeautifulSoup object from html in zip file"""
    html_doc = extract_html(file_name)
    soup = BeautifulSoup(html_doc)
    return soup

def get_body_soup(file_name):
    """Returns the body of the html from the zip file"""
    soup = get_html_soup(file_name)
    body = soup.find("body")
    return body

def get_css_text(file_name):
    """Returns the text of the document style"""
    soup = get_html_soup(file_name)
    style = soup.find("style")
    return style.contents[0]

def get_tinycss_sheet(file_name):
    """Returns a tinycss sheet object from the zip file"""
    parser = tinycss.css21.CSS21Parser()
    css_text = get_css_text(file_name)
    sheet = parser.parse_stylesheet(css_text)
    return sheet

def get_inline_style(tinycss_sheet, tag):
    """Returns the inline style string for a css class, tag, etc"""
    rule = None
    for r in tinycss_sheet.rules:
        if r.selector.as_css() == tag:
            rule = r
    output = []
    for d in rule.declarations:
        output.append(d.name + ":" + " ".join([x.as_css() for x in d.value]))
    return ";".join(output)

def clear_tab_attr(soup, tag):
    """Clears all attributes for all tags of a certain type in a BeautifulSoup object"""
    tagged = soup.findAll(tag)
    for tag in tagged:
        tag.attrs = None

def combine_tags(parent, tag):
    """Combines consecutive objects of the same tag into a single tag"""
    # Clarification: Google Docs allows you to style paragraphs as
    #   specific header types. However, Google Docs handles line-
    #   breaks as endtag/newtag, rather than <br/>, so if you have 
    #   borders, padding, etc, on one of those tags in your css, 
    #   they will appear on every line break. This function collapses
    #   consecutive tags into one.
    #   TL:DR; <h6>Herp</h6><h6>Derp</h6> becomes <h6>Herp<br/>Derp</h6>
    last_child = None
    for child in parent.children:
        name = child.name
        if name == "hr":
            combine_tags(child, tag)
        elif not name == tag:
            last_child = child
        elif name == tag:
            if last_child is None:
                last_child = child
            elif last_child.name != tag:
                last_child = child
            else:
                last_child.contents.append(BeautifulSoup("<br/>"))
                last_child.contents += child.contents
                child.clear()
    must_clear = parent.findAll(tag)
    must_clear.reverse()
    for tag in must_clear:
        if len(tag.findChildren()) == 0:
            tag.extract()

def split_pages(parent):
    """Returns a list of pages from a soup object"""
    # Google docs splits pages by embedding them in 
    #   <hr> tags.
    output = []
    hr = parent.find("hr")
    new_page = None
    if hr:
        new_page = hr.extract()
    output.append(parent)
    if new_page:
        output += split_pages(hr)
    return output

def strip_attr(tag, attr_name):
    """Remove an attribute from a tag"""
    attrs = tag.attrs
    new_attrs = {}
    for k in attrs.keys():
        if not k == attr_name:
            new_attrs[k] = attrs[k]
    tag.attrs = new_attrs

def strip_soup_tags_attr(soup, tag, attr_name):
    """Remove an attribute from all of a certain tag in a BeautifulSoup object"""
    tags = soup.findAll(tag)
    for t in tags:
        strip_attr(t, attr_name)

def replace_tag_with_inline_css(styles, tag):
    """Replace all of a certain tag with the inline CSS style"""
    # Why inline styles? Because in the middle of a paragraph, 
    #   if you decided to make something bold or italic or a 
    #   slightly different font or color, there's no sense in 
    #   dumping it into your master file.
    class_ = tag.get("class")
    if class_:
        strip_attr(tag, "style")
        strip_attr(tag, "class")
        new_styles = []
        for c in class_:
            style_string = get_inline_style(styles, "." + c)
            new_styles.append(style_string)
        tag.attrs["style"] = ";".join(new_styles)

def strip_class_from_basics(soup):
    """Remove the class from all of certain tags"""
    # Odds are, you have CSS for the following ready to go
    basics = ["title", "subtitle", "h1", "h2", "h3", "h4", "h5", "h6", "a", "p"]
    for b in basics:
        tags = soup.findAll(b)
        for tag in tags:
            strip_attr(tag, "class")
            
def change_class_to_inlines(soup, styles):
    """Change class to inline style for all of certain tags"""
    use_inline = ["span",]
    for b in use_inline:
        tags = soup.findAll(b)
        for tag in tags:
            replace_tag_with_inline_css(styles, tag)

def unpack_zip_to_html(file_name, output_dir):
    """Processes html and image files from GoogleDoc output and places the resulting files in output_dir"""
    fn = file_name
    new_dir = os.path.join(output_dir, fn.split("\\")[-1].replace(".zip", ""))
    os.mkdir(new_dir)
    body = get_body_soup(fn)
    sheet = get_tinycss_sheet(fn)
    strip_class_from_basics(body)
    change_class_to_inlines(body, sheet)
    combine_tags(body, "h6")
    pages = split_pages(body)
    for i in range(len(pages)):
        f = os.path.join(new_dir, "page_{0}.html".format(str(i)))
        f_out = open(f, "w")
        byte_count = f_out.write(pages[i].prettify(formatter="html"))
        f_out.close()
    extract_pngs(file_name, new_dir)
