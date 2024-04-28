import os.path
from random import randint

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

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


def populate_slide(presentation_id, page_id, page_type, text_box_text):
    try:
        service = build("slides", "v1", credentials=getCreds())

        shape_id_list = []
        pt350 = {"magnitude": 350, "unit": "PT"}
        # TODO https://developers.google.com/slides/api/guides/styling
        requests = [
            {
                "createShape": {
                    "objectId": generate_id("shape", shape_id_list=shape_id_list),
                    "shapeType": "TEXT_BOX",
                    "elementProperties": {
                        "pageObjectId": page_id,
                        "size": {"height": pt350, "width": pt350},
                        "transform": {
                            "scaleX": 1,
                            "scaleY": 1,
                            "translateX": 350,
                            "translateY": 100,
                            "unit": "PT",
                        },
                    },
                }
            },
            {
                "insertText": {
                    "objectId": shape_id_list[0],
                    "insertionIndex": 0,
                    "text": text_box_text,
                }
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


def create_slide(presentationId, type, body_text):
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
        populate_slide(presentationId, create_slide_response.get('objectId'), type, body_text)

    except HttpError as error:
        print(f"An error occured: {error}")
        print("Slides not created")
        return error


if __name__ == "__main__":
    # Put the title of the presentation
    adminPresentation = create_presentation("[TEST] Diego Presentation")
    create_slide(adminPresentation.get('presentationId'), "CALL_TO_WORSHIP", "In the early church, Lent was a time of preparation for baptism at Easter.")
