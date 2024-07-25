import os.path
from random import randint

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from ExtractFromPDF import ExtractorClass

SCOPES = ['https://www.googleapis.com/auth/presentations']


def getCreds():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    return creds


def create_presentation(title):
    # pylint: disable=maybe-no-member
    try:
        service = build("slides", "v1", credentials=getCreds())

        body = {"title": title}
        presentation = service.presentations().create(body=body).execute()
        print(
            f"Created presentation with ID:{(presentation.get('presentationId'))}"
        )
        return presentation

    except HttpError as error:
        print(f"An error occurred: {error}")
        print("presentation not created")
        return error


def generate_id(element_type, _id='', shape_id_list=None, page_id_list=None):
    for i in range(8):
        _id = _id + str(randint(0, 9))
    if element_type == "shape" and _id not in shape_id_list:
        shape_id_list.append(_id)
    elif element_type == "shape":
        _id = generate_id(element_type, shape_id_list)
    if element_type == "page" and _id not in page_id_list:
        page_id_list.append(_id)
    elif element_type == "page":
        _id = generate_id(element_type)
    return _id


slide_types = {
    "CALL_TO_WORSHIP": []
}


def populate_slide(presentation_id, page_id, slide_title, text_box_text, section_only):
    try:
        service = build("slides", "v1", credentials=getCreds())

        shape_id_list = []
        pt405 = {"magnitude": 405, "unit": "PT"}
        cm1_62 = {"magnitude": cm_to_pt(1.62), "unit": "PT"}
        pt720 = {"magnitude": 720, "unit": "PT"}
        # TODO https://developers.google.com/slides/api/guides/styling
        if not section_only:
            requests = [
                {
                    "createShape": {
                        "objectId": generate_id("shape", shape_id_list=shape_id_list),
                        "shapeType": "TEXT_BOX",
                        "elementProperties": {
                            "pageObjectId": page_id,
                            "size": {"height": pt405, "width": pt720},
                            # "transform": {
                            #     "scaleX": 1,
                            #     "scaleY": 1,
                            #     "translateX": 350,
                            #     "translateY": 100,
                            #     "unit": "PT",
                            # },
                        },
                    }
                },
                {
                    "updateShapeProperties": {
                        "objectId": shape_id_list[0],
                        "shapeProperties": {
                            "contentAlignment": "TOP"
                        },
                        "fields": "contentAlignment",
                    }
                },
                {
                    "insertText": {
                        "objectId": shape_id_list[0],
                        "insertionIndex": 0,
                        "text": text_box_text,
                    }
                },
                {  # https://developers.google.com/slides/api/reference/rest/v1/presentations.pages/text#Page.TextStyle
                    "updateTextStyle": {
                        "objectId": shape_id_list[0],
                        "fields": "*",
                        "style": {
                            "bold": True,
                            "underline": True,
                            "foregroundColor": {
                                "opaqueColor": {
                                    "rgbColor": {
                                        "red": 1.0,
                                        "green": 1.0,
                                        "blue": 1.0
                                    }
                                }
                            },
                            "fontFamily": "nunito",
                            "fontSize": {
                                "magnitude": "22",
                                "unit": "PT"
                            }
                        },
                        "textRange": {"type": "ALL"}
                    },
                },
                # {
                #     "insertText": {
                #         "objectId": shape_id_list[0],
                #         "insertionIndex": 0,
                #         "text": slide_title,
                #     }
                # },
            ]
        else:
            requests = [
                {
                    "createShape": {
                        "objectId": generate_id("shape", shape_id_list=shape_id_list),
                        "shapeType": "TEXT_BOX",
                        "elementProperties": {
                            "pageObjectId": page_id,
                            "size": {"height": cm1_62, "width": pt720},
                            "transform": {
                                "scaleX": 1,
                                "scaleY": 1,
                                "translateX": 0,
                                "translateY": cm_to_pt(9.58),
                                "unit": "PT",
                            },
                        },
                    }
                },
                {
                    "updateShapeProperties": {
                        "objectId": shape_id_list[0],
                        "shapeProperties": {
                            "contentAlignment": "BOTTOM"
                        },
                        "fields": "contentAlignment",
                    }
                },
                {
                    "insertText": {
                        "objectId": shape_id_list[0],
                        "insertionIndex": 0,
                        "text": slide_title,
                    }
                },
                {  # https://developers.google.com/slides/api/reference/rest/v1/presentations.pages/text#Page.TextStyle
                    "updateTextStyle": {
                        "objectId": shape_id_list[0],
                        "fields": "*",
                        "style": {
                            "bold": True,
                            "foregroundColor": {
                                "opaqueColor": {
                                    "rgbColor": {
                                        "red": 1.0,
                                        "green": 1.0,
                                        "blue": 1.0
                                    }
                                }
                            },
                            "fontFamily": "oswald",
                            "fontSize": {
                                "magnitude": "52",
                                "unit": "PT"
                            },
                        },
                        "textRange": {"type": "ALL"}
                    },
                },
            ]
        body = {"requests": requests}
        response = (
            service.presentations()
                .batchUpdate(presentationId=presentation_id, body=body)
                .execute()
        )
        create_shape_response = response.get("replies")[0].get("createShape")
        print(f"Created textbox with ID:{(create_shape_response.get('objectId'))}")
    except HttpError as error:
        print(f"An error occured: {error}")
        return error
    return response


def create_slide(presentationId, slide_title, body_text, is_section_only=False):
    page_id_list = []
    try:
        service = build("slides", "v1", credentials=getCreds())
        requests = [
            {
                "createSlide": {
                    "objectId": generate_id(element_type="page", page_id_list=page_id_list),
                    "slideLayoutReference": {
                        "predefinedLayout": "BLANK"
                    },
                }
            },
            {
                "updatePageProperties": {
                    "objectId": page_id_list[0],
                    "pageProperties": {
                        "pageBackgroundFill": {
                            "stretchedPictureFill": {
                                "contentUrl": "https://lh3.google.com/u/0/d/1nxx9KSVqdvFcZKiaNG0UcousOFtXOLKj=w2560-h1282-iv2"
                            }
                        }
                    },
                    "fields": "pageBackgroundFill"
                }
            }
        ]

        body = {'requests': requests}
        response = (
            service.presentations()
                .batchUpdate(presentationId=presentationId, body=body)
                .execute())
        create_slide_response = response.get("replies")[0].get("createSlide")
        print(f"Created slide with ID:{(create_slide_response.get('objectId'))}")
        populate_slide(presentationId, create_slide_response.get('objectId'), slide_title, body_text,
                       section_only=is_section_only)

    except HttpError as error:
        print(f"An error occurred: {error}")
        print("Slides not created")
        return error


def cm_to_pt(cm):
    return (cm/2.54)/(1/72)


if __name__ == "__main__":
    adminPresentation = create_presentation("[TEST] Diego Presentation")
    pres_id = adminPresentation.get("presentationId")
    extractor = ExtractorClass()
    sections = extractor.identify_sections()
    section_contents = extractor.match_content_with_section(sections)

    for section in sections:
        if not section_contents[sections.index(section)] == "No Content":
            create_slide(pres_id, section + "\n", section_contents[sections.index(section)])
        else:
            create_slide(pres_id, section, section_contents[sections.index(section)], is_section_only=True)
        # print(section + "\n" + section_contents[sections.index(section)] + "\n\n\n")
    # TODO When using section names in CreateSlideshow.py remove * from string
