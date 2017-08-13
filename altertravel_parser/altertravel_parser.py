"""
This tool parses http://altertravel.ru for the Geo Points and generates KML file.
"""
import argparse
import os
import re
import sys
import urllib
import urllib2

import progressbar
from bs4 import BeautifulSoup
from lxml import etree
from pykml import parser
from pykml.factory import KML_ElementMaker as Kml

SITE = 'http://altertravel.ru'

arguments = argparse.ArgumentParser()
arguments.add_argument('--tag', type=lambda s: unicode(s, sys.getfilesystemencoding()),
                       help='tag on the http://altertravel.ru',
                       required=True)
arguments = arguments.parse_args()

# Loading results for the search by the tag
print('Parsing pages for the results')
tag = arguments.tag.encode('utf-8')
sights_ids = []
page = 0
while True:
    print('Page {}'.format(page))
    url = '{}/catalog_sub.php?p={}&filter=&stat=&d=&tag={}'.format(SITE, page, urllib.quote(tag))
    content = urllib2.urlopen(url).read()
    parsed_html = BeautifulSoup(content, 'html.parser')
    sights = parsed_html.find_all('div', attrs={'class': 'info_title'})
    if not sights:
        break

    for sight in sights:
        sight_id = re.search('id=([0-9]*)', sight.a.get('href')).group(1)
        sights_ids.append(sight_id)
    page += 1

cache = []
if os.path.exists('cache'):
    with open('cache') as cache_file:
        cache = cache_file.read().splitlines()

kml_folder = Kml.Folder(Kml.name(tag.decode('utf-8-sig')))
sights_count = len(sights_ids)
print('Parsing {} KMLs'.format(sights_count))
bar = progressbar.ProgressBar(maxval=sights_count,
                              widgets=[progressbar.Bar('=', '[', ']'), ' ', progressbar.Percentage()])
bar.start()
new_point_count = 0
for index, sight_id in enumerate(sights_ids):
    # Loading and parsing KMLs
    kml_str = urllib2.urlopen('{}/generate_kml.php?id={}'.format(SITE, sight_id))
    root = parser.parse(kml_str).getroot()
    point_name = root.Document.Placemark.name.text
    point_coordinates = root.Document.Placemark.Point.coordinates.text
    bar.update(index)

    if point_coordinates in cache:
        continue

    # Loading description for the point
    content = urllib2.urlopen('{}/view.php?id={}'.format(SITE, sight_id)).read()
    parsed_html = BeautifulSoup(content, 'html.parser')
    description = parsed_html.find('div', attrs={'class': 'post col-sm-12'})
    for div in description.find_all('div'):
        div.extract()
    description = Kml.description(description.p.text.strip() if description.p else '')
    bar.update(index + 1)
    point = Kml.Placemark(Kml.name(point_name), Kml.Point(Kml.coordinates(point_coordinates)), description)
    kml_folder.append(point)
    cache.append(point_coordinates)
    new_point_count += 1

bar.finish()

print('{} new points'.format(new_point_count))

with open('cache', 'wt') as cache_file:
    cache_file.write('\n'.join(cache))

kml = Kml.kml(kml_folder)
filename = arguments.tag.encode(sys.getfilesystemencoding())
with open('{}.kml'.format(filename), 'wt') as kml_file:
    kml_file.write(etree.tostring(kml, pretty_print=True))
