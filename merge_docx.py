import docx
from docx import Document
from docx.oxml.styles import CT_Style
import os
import random


# sub 문서의 모든 인라인 이미지를 template에 추가하는 함수
def handle_inlines(template, sub):
    x_path = '//a:blip'
    # sub 문서의 모든 a:blip 엘리먼트를 blip_list 변수에 저장
    blip_list = sub.element.xpath(x_path)

    # shapes 변수에는 관심 있는 미디어를 가진 모든 InlineShape 객체가 들어 있음
    shapes = sub.inline_shapes

    # 각 이미지를 고유한 파일 경로로 저장
    for i in range(len(shapes)):

        # rId를 가져오기
        shape = shapes[i]
        if shape._inline.graphic.graphicData.pic is not None:
            rId = shape._inline.graphic.graphicData.pic.blipFill.blip.embed
        else:
            # 해당 shape를 제거
            drawing_element = shape._inline.getparent()
            drawing_parent = drawing_element.getparent()
            drawing_parent.remove(drawing_element)
            continue

        # ImagePart 객체
        image_part = sub.part.related_parts[rId]

        # Image 객체
        actual_image = image_part.image

        # 이미지의 실제 바이너리 데이터
        image_blob = actual_image.blob

        # 고유한 파일 경로 문자열 생성
        random_id = random.randint(10000, 100000)
        image_path = 'image' + str(random_id) + '.' + actual_image.ext

        # 이미지의 바이너리 데이터를 이미지 파일로 저장
        image_file = open(image_path, "wb")
        image_file.write(image_blob)
        image_file.close()

        # template 파일에 sub_doc의 미디어에 대한 새로운 Relationship 생성
        new_rId, img = template.part.get_or_add_image(image_path)
        print(type(template.part), type(img))
        # new_rId는 추가된 이미지의 새로운 rId를 가리킴
        # img는 Image 객체를 가리킴(사용 안 함)

        blip_element = blip_list[i]
        # blip_list의 각 요소에 대해 새 rId 값을 설정
        blip_element.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed', new_rId)

        # 이미지 파일 제거
        os.remove(image_path)


# sub 문서의 모든 스타일을 template의 'styles.xml' 파일에 추가하는 함수
# 이미 존재하는 스타일은 덮어써짐
def handle_styles(template, sub):
    # 템플릿 파일의 styles 엘리먼트를 저장하는 변수
    template_styles = template.styles.element

    # sub 문서의 styles 엘리먼트를 저장하는 변수
    sub_styles = sub.styles.element

    # sub 문서의 각 스타일을 확인하여 CT_Style 객체인 경우 'styles.xml' 파일에 추가
    for s in sub_styles:
        if isinstance(s, CT_Style):
            template_styles.append(s)


def merge_docx(file_list, file_name):
    merged_doc = Document()

    for file in file_list:
        sub_doc = Document(file)
        # 인라인 이미지 처리
        handle_inlines(merged_doc, sub_doc)
        # 스타일 처리
        handle_styles(merged_doc, sub_doc)

        # sub_doc의 body 엘리먼트를 template 파일의 body에 추가
        for element in sub_doc.element.body:
            merged_doc.element.body.append(element)

    # 문서 저장
    merged_doc.save(file_name)


file_list = ['test1.docx', 'test2.docx']
merge_docx(file_list, 'output.docx')