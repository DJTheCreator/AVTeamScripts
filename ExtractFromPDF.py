from pypdf import PdfReader


class ExtractorClass:
    def __init__(self):
        reader = PdfReader("sunday.pdf")
        _bulletin = ''
        for page in reader.pages:
            _bulletin += page.extract_text()
        self.bulletin = _bulletin

    def identify_sections(self):
        section_list = []
        for line in self.bulletin.splitlines():
            for word in line.split():
                if word.isupper() and word != 'PLEASE' and not line.startswith('  '):
                    if 'CHIMES' in line or 'CALL TO WORSHIP' in line:  # TODO consider asking user for name of first section
                        section_list.clear()
                    section_list.append(line)
                    break
                else:
                    continue
        for line in section_list:
            for char in line:
                try:
                    end_index = line.index('   ')
                    line_index = section_list.index(line)
                    section_list[line_index] = line[:end_index].lstrip()
                except ValueError:
                    pass
        return section_list

    def match_content_with_section(self, section_list: [str]):
        section_content_list = []
        for section in section_list:
            for line in self.bulletin.splitlines():
                if section in line:
                    next_line_index = self.bulletin.splitlines().index(line) + 1
                    break
            current_section_index = self.bulletin.index(
                self.bulletin.splitlines()[next_line_index])  # the index of the start of the next line in bulletin
            try:
                next_section_index = self.bulletin.index(section_list[section_list.index(section) + 1])
                section_content = self.bulletin[current_section_index:next_section_index]
            except IndexError:
                section_content = self.bulletin[current_section_index:]
            if self.bulletin[current_section_index] == ' ':  # if next line is blank (current section has no content)
                section_content = 'No Content'
            section_content = self.clean_sections(section_content)
            section_content_list.append(section_content)
        return section_content_list

    def clean_sections(self, content):
        content = ''.join(content.splitlines())
        content = ' '.join(content.split())
        for word in content.split():
            if 'pastor' in word.lower():
                content = content.replace(word, '\n/p')
            if 'leader' in word.lower():
                content = content.replace(word, '\n/p')
            if 'people' in word.lower():
                content = content.replace(word, '\n/b')
        content = content.lstrip()
        return content


extractor = ExtractorClass()
sections = extractor.identify_sections()
section_contents = extractor.match_content_with_section(sections)
# print(section_contents[sections.index(sections[2])])
# for section in sections:
#     print(section + "\n" + section_contents[sections.index(section)] + "\n\n\n")
# # TODO When using section names in CreateSlideshow.py remove * from string
