# !/usr/bin/python
import sys
import cgi
import urllib
import glob
import os
path = os.getcwd()

# you will need python-docx for this.
import docx

# you will need csv for this
import csv

# you will need xmltodict for this
import xmltodict

# API CODE FOR THE COPYSCAPE

if sys.hexversion < 0x02050000:
    import \
        elementtree.ElementTree as CopyscapeTree
else:
    import xml.etree.ElementTree as CopyscapeTree
isPython2 = sys.hexversion < 0x03000000
if isPython2:
    import urllib2
else:
    import urllib.request
    import urllib.error


# API SETTINGS
COPYSCAPE_USERNAME = ""
COPYSCAPE_API_KEY = ""
COPYSCAPE_API_URL = "http://www.copyscape.com/api/"


def copyscape_api_url_search_internet(url, full=0):
    return copyscape_api_url_search(url, full, 'csearch')

def copyscape_api_text_search_internet(text, encoding, full=0):
    return copyscape_api_text_search(text, encoding, full, 'csearch')

def copyscape_api_check_balance():
    return copyscape_api_call('balance')

def copyscape_api_url_search_private(url, full=0):
    return copyscape_api_url_search(url, full, 'psearch')

def copyscape_api_url_search_internet_and_private(url, full=0):
    return copyscape_api_url_search(url, full, 'cpsearch')

def copyscape_api_text_search_private(text, encoding, full=0):
    return copyscape_api_text_search(text, encoding, full, 'psearch')

def copyscape_api_text_search_internet_and_private(text, encoding, full=0):
    return copyscape_api_text_search(text, encoding, full, 'cpsearch')

def copyscape_api_url_add_to_private(url, id=None):
    params = {}
    params['q'] = url
    if id is not None:
        params['i'] = id

    return copyscape_api_call('pindexadd', params)

def copyscape_api_text_add_to_private(text, encoding, title=None, id=None):
    params = {}
    params['e'] = encoding
    if title is not None:
        params['a'] = title
    if id != None:
        params['i'] = id

    return copyscape_api_call('pindexadd', params, text)


def copyscape_api_delete_from_private(handle):
    params = {}
    if handle is None:
        params['h'] = ''
    else:
        params['h'] = handle

    return copyscape_api_call('pindexdel', params)


def copyscape_api_url_search(url, full=0, operation='csearch'):
    params = {}
    params['q'] = url
    params['c'] = str(full)

    return copyscape_api_call(operation, params)


def copyscape_api_text_search(text, encoding, full=0, operation='csearch'):
    params = {}
    params['e'] = encoding
    params['c'] = str(full)

    return copyscape_api_call(operation, params, text)

# REMOVE THE LINE WHERE X PARAMETER IS MADE 1 FOR LIVE VERSION.

def copyscape_api_call(operation, params={}, postdata=None):
    urlparams = {}
    urlparams['u'] = COPYSCAPE_USERNAME
    urlparams['k'] = COPYSCAPE_API_KEY
    urlparams['o'] = operation
    urlparams['x'] = 1
    urlparams.update(params)

    uri = COPYSCAPE_API_URL + '?'

    request = None
    if isPython2:
        uri += urllib.urlencode(urlparams)
        if postdata is None:
            request = urllib2.Request(uri)
        else:
            request = urllib2.Request(uri, postdata.encode("UTF-8"))
    else:
        uri += urllib.parse.urlencode(urlparams)
        if postdata is None:
            request = urllib.request.Request(uri)
        else:
            request = urllib.request.Request(uri, postdata.encode("UTF-8"))

    try:
        response = None
        if isPython2:
            response = urllib2.urlopen(request)
        else:
            response = urllib.request.urlopen(request)
        return response

    except Exception:
        e = sys.exc_info()[1]
        print(e.args[0])

    return None


def copyscape_title_wrap(title):
    return title + ":"


def copyscape_node_wrap(element):
    return copyscape_node_recurse(element)


def copyscape_node_recurse(element, depth=0):
    ret = ""
    if element is None:
        return ret

    ret += "\t" * depth + " " + element.tag + ": "
    if element.text is not None:
        ret += element.text.strip()
    ret += "\n"
    for child in element:
        ret += copyscape_node_recurse(child, depth + 1)

    return ret

def keys_exists(element, *keys):
    '''
    Check if *keys (nested) exists in `element` (dict).
    '''
    if not isinstance(element, dict):
        raise AttributeError('keys_exists() expects dict as first argument.')
    if len(keys) == 0:
        raise AttributeError('keys_exists() expects at least two arguments, one given.')

    _element = element
    for key in keys:
        try:
            _element = _element[key]
        except KeyError:
            return False
    return True

# RUNNING THE SCRIPT TO READ THE FILES AND GENERATE A SINGLE FILE FOR THE RESULTS
directory = path+"/files/"
generated = path+"/generated/"
list = glob.glob(directory+"*.docx")

# File name for generating
filename = "aggregatedfile"
with open(generated + filename + '.csv', 'w', newline='') as file:
    writer = csv.writer(file)
    writer.writerow(
        ["querywords", "cost", "count", "index", "url", "title", "textsnippet", "minwordsmatched", "viewurl",
         "urlwords", "wordsmatched", "textmatched", "percentmatched"])

    for x in list:

        #reset for the document
        completeText = ''
        result = ''

        #reading the document
        doc = docx.Document(x)
        all_paras = doc.paragraphs

        #combining all paras
        for para in all_paras:
            completeText = completeText + ' ' + para.text

        #using the api to fetch the records against the file
        result = copyscape_api_text_search_internet(completeText, "UTF-8", 2)

        #parse the result from XML
        res = result.read()
        doc = xmltodict.parse(res)
        querywords = doc['response']['querywords'];
        cost = doc['response']['cost'];
        count = doc['response']['count'];
        result = doc['response']['result'];

        print(result)

        #writing the file
        writer = csv.writer(file)
        for row in result:

              if(keys_exists(row, "index")):
                  index = row['index']
              else:
                  index = ' '

              if(keys_exists(row, "url")):
                  url = row['url']
              else:
                  url = ' '

              if(keys_exists(row, "title")):
                  title = row['title']
              else:
                  title = ' '

              if(keys_exists(row, "textsnippet")):
                  textsnippet = row['textsnippet']
              else:
                  textsnippet = ' '

              if(keys_exists(row, "minwordsmatched")):
                  minwordsmatched = row['minwordsmatched']
              else:
                  minwordsmatched = ' '

              if(keys_exists(row, "viewurl")):
                  viewurl = row['viewurl']
              else:
                  viewurl = ' '

              if(keys_exists(row, "urlwords")):
                  urlwords = row['urlwords']
              else:
                  urlwords = ' '

              if(keys_exists(row, "wordsmatched")):
                  wordsmatched = row['wordsmatched']
              else:
                  wordsmatched = ' '

              if(keys_exists(row, "textmatched")):
                  textmatched = row['textmatched']
              else:
                  textmatched = ' '

              if(keys_exists(row, "viewurl")):
                  viewurl = row['viewurl']
              else:
                  viewurl = ' '

              if(keys_exists(row, "percentmatched")):
                  percentmatched = row['percentmatched']
              else:
                  percentmatched = ' '


              writer.writerow([querywords, cost, count,index,url,title,textsnippet.encode("utf-8"),minwordsmatched,viewurl,urlwords,wordsmatched,textmatched.encode("utf-8"),percentmatched])



