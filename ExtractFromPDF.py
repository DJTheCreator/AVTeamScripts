from pypdf import PdfReader

reader = PdfReader("sunday.pdf")
bulletin = ''
for page in reader.pages:
    bulletin += page.extract_text()


# TODO Search for 'People:' and find index where people lines end to have in and out points for bold
# for i in range(len(bulletin.split())):
#     if bulletin.split()[i].lower() == "gathering":
#         bulletin = bulletin.split()[i+1:]
#         break
# print()

def identify_sections():
    section_list = []
    for line in bulletin.splitlines():
        for word in line.split():
            if word.isupper() and word != 'PLEASE' and not line.startswith('  '):
                if 'CHIMES' in line or 'CALL TO WORSHIP' in line:  # TODO consider asking user for name of first section
                    section_list.clear()
                section_list.append(line)
                break
            else:
                continue
    for line in section_list:
        count = 0
        for char in line:
            try:
                end_index = line.index('   ')
                line_index = section_list.index(line)
                section_list[line_index] = line[:end_index].lstrip()
            except ValueError:
                pass
    return section_list


def match_content_with_section(section_list):
    section_content_list = []
    for section in section_list:
        for line in bulletin.splitlines():
            if section in line:
                next_line_index = bulletin.splitlines().index(line) + 1
                break
        current_section_index = bulletin.index(
            bulletin.splitlines()[next_line_index])  # the index of the start of the next line in bulletin
        try:
            next_section_index = bulletin.index(section_list[section_list.index(section) + 1])
            # print("In section: " + section + "\nstart index: " + str(current_section_index) + "\nending index: " + str(next_section_index))
            section_content = bulletin[current_section_index:next_section_index]
        except IndexError:
            section_content = bulletin[current_section_index:]
        if bulletin[current_section_index] == ' ':  # if next line is blank (current section has no content)
            section_content = 'No Content'
        section_content_list.append(section_content)
    return section_content_list


sections = identify_sections()
section_contents = match_content_with_section(sections)
for section in sections:
    print(section + "\n" + section_contents[sections.index(section)] + "\n\n\n")
# TODO When using section names in CreateSlideshow.py remove * from string
