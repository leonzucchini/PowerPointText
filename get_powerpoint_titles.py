'''Access content of a pptx document and copy file titles to a text file. '''

import sys
import os
import re
import zipfile

def get_user_input():
    '''Get user input from command line and return file paths. '''

    if len(sys.argv) < 2:
        print 'usage: ./get_powerpoint_titles.py file-to-read [--file-to-write]'
        sys.exit(1)

    else:
        file_to_read = sys.argv[1]

        if len(sys.argv) == 2:
            result_folder = os.path.dirname(os.path.abspath(__file__))
            file_to_write = os.path.join(result_folder, 'pptx_titles.txt')
        else:
            file_to_write = sys.argv[2]

    return (file_to_read, file_to_write)

def get_zip(file_path):
    '''Get file and convert to zip file. '''

    zip_file = zipfile.ZipFile(file_path)
    return zip_file

def get_slide_names(zip_file):
    '''Get names of all slides in the file, return list of names. '''

    slide_tuples = []
    slides = []

    slide_names = zip_file.namelist()
    pattern = re.compile(r'ppt/slides/slide\d+.*')

    # Extract slides names from files using regex
    for slide_name in slide_names:
        match = re.match(pattern, slide_name)
        if match:
            name = match.group()
            order = int(re.match(r'.*slide(\d+).xml', name).group(1))
            slide_tuples.append((name, order))

    # Sort slides
    slide_tuples = sorted(slide_tuples, key=lambda slide: slide[1])
    for slide in slide_tuples:
        slides.append(slide[0])

    return slides

def get_title(zip_file, slide_name):
    '''Get title of slide. '''

    xml = zip_file.read(slide_name)
    pattern = re.compile(r'<p:cSld><p:spTree>.*?<p:sp>.*?<p:txBody>.*?<a:p><a:r>.*?<a:t>(.*?)</a:t>')
    match = re.findall(pattern, xml)

    if match:
        return match[0]
    else:
        return None

def get_titles(zip_file, slide_names):
    '''Get all titles of the presentation. '''

    titles = []
    for slide_name in slide_names:
        title = get_title(zip_file, slide_name)
        if title:
            titles.append(title)

    return titles

def main():
    '''Get user input on files, extract and store slide titles. '''
    # Get user input
    user_input = get_user_input()
    file_to_read = user_input[0]
    file_to_write = user_input[1]

    # Get slide tiles
    zip_file = get_zip(file_to_read)
    slide_names = get_slide_names(zip_file)
    titles = get_titles(zip_file, slide_names)

    # Write to file
    with open(file_to_write, 'w') as f:
        titles = '\n'.join(titles)
        f.write(titles)

    print 'Extracted titles from %s and wrote them to %s' %(file_to_read, file_to_write)

if __name__ == '__main__':
    main()

# P.S. Yes I know there are easier ways to do this in pptx - this was just for fun :)
