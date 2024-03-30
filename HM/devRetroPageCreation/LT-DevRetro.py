from requests.auth import HTTPBasicAuth
from datetime import datetime, timedelta
import requests
import json

# Hayanmind Notion API Token
API_KEY = '노션 API 키'

headers = {
    "Authorization": "Bearer " + API_KEY,
    "Content-Type": "application/json",
    "Notion-Version": "2022-06-28"
}

members_notion = ['노션 닉네임']
members_Jira = ['지라 닉네임']
created_date = datetime.now().strftime('%Y-%m-%d')
meeting_date = (datetime.now() + timedelta(days=3)).strftime('%Y-%m-%d')

## Returns list of issues of an assignee from Jira
def getIssuesFromJira(member_jira):
    url = "https://hayanmind.atlassian.net/rest/api/3/search"
    auth = HTTPBasicAuth('아이디', '인증 키 값')
    headers = {
        "Accept": "application/json"
    }

    issues = []
    sprint = ''

    query = {
        'jql': f'project in ("[프로젝트 명]") AND Sprint in openSprints() AND assignee = "{member_jira}"'
    }

    response = requests.request(
        "GET",
        url,
        headers=headers,
        params=query,
        auth=auth
    )

    json_data = response.json()

    for key in json_data['issues']:
        if (key['fields']['status']['name'] == "NO ISSUE"):
            continue
        issues.append(key['key'])

        if (key['fields']['customfield_10020'][0]['state'] == 'active'):
            sprint = key['fields']['customfield_10020'][0]['name']

    return issues, sprint[-2:]

## Creates meeting page on Notion
## this returns ID of created page
def createPage(headers, target):
    createUrl = 'https://api.notion.com/v1/pages'
    issues, sprintNum = getIssuesFromJira("첫번째 이름")

    with open(f'{target}') as f:
        json_data = json.load(f)

    json_data['properties']['Date']['date']['start'] = f'{meeting_date} 10:00'

    json_data['properties']['Title']['title'][0]['text']['content'] = f'NX Sprint {sprintNum} Dev Retrospective'
    data = json.dumps(json_data, sort_keys=True, indent=2)
    res = requests.request("POST", createUrl, headers=headers, data=data)

    json_data = res.json()

    return json_data["url"][-32:]

# This returns the latest block of newely created page
def getTargetBlock(pageID):
    createUrl = f'https://api.notion.com/v1/blocks/{pageID}/children?page_size=100'
    targetID = ''
    headers = {
        "Authorization": "Bearer " + API_KEY,
        "accept": "application/json",
        "Notion-Version": "2022-06-28"
    }

    res = requests.request("GET", createUrl, headers=headers)
    json_data = res.json()
    for data in json_data["results"]:
        targetID = data['id']

    return targetID

# This writes data from JIRA for each members
def createPageContent(blockID, members_notion, member_index, isLast):
    issues = []
    sprintNum = 0

    json_data = {
        'children': []
    }
    createUrl = f'https://api.notion.com/v1/blocks/{blockID}/children'
    issues, sprintNum = getIssuesFromJira(members_Jira[member_index])
    if isLast == False:
        json_data['children'].append({
            "type": "heading_1",
            "heading_1": {
                "rich_text": [{
                    "type": "text",
                    "text": {
                        "content": f'{members_notion}'
                    }
                }]
            },
        })
        json_data['children'].append({
            "type": "heading_2",
            "heading_2": {
                "rich_text": [{
                    "type": "text",
                    "text": {
                        "content": "Work Done"
                    }
                }],
                "color": "gray_background"
            }
        })
        json_data['children'].append({
            "type": "paragraph",
            "paragraph": {
                "rich_text": [{
                    "type": "text",
                    "text": {
                        "content": "📝 What tasks did you get done during the last sprint? (Paste links to Issues on Jira)"
                    },
                    "annotations": {
                        "color": "gray",
                    }
                }]
            }
        })
        for issue in issues:
            json_data['children'].append({
                "type": "paragraph",
                "paragraph": {
                    "rich_text": [
                        {
                            "type": "text",
                            "text": {
                                "content": f'https://hayanmind.atlassian.net/browse/{issue}',
                                "link": {
                                    "type": "url",
                                    "url": f'https://hayanmind.atlassian.net/browse/{issue}'
                                }
                            }
                        }
                    ]
                }
            })
        json_data['children'].append({
            "type": "heading_2",
            "heading_2": {
                "rich_text": [{
                    "type": "text",
                    "text": {
                        "content": "Retrospective"
                    }
                }],
                "color": "gray_background"
            }
        })
        json_data['children'].append({
            "type": "paragraph",
            "paragraph": {
                "rich_text": [{
                    "type": "text",
                    "text": {
                        "content": "🔵 What are some things you liked about or think you did well during the last sprint?"
                    },
                    "annotations": {
                        "color": "gray",
                    }
                }]
            }
        })
        json_data['children'].append({
            "type": "bulleted_list_item",
            "bulleted_list_item": {
                "rich_text": [{
                    "type": "text",
                    "text": {
                        "content": ""
                    }
                }]
            }
        })
        json_data['children'].append({
            "type": "paragraph",
            "paragraph": {
                "rich_text": [{
                    "type": "text",
                    "text": {
                        "content": "🔴 What are some things you felt lacking or think could have gone better?"
                    },
                    "annotations": {
                        "color": "gray",
                    }
                }]
            }
        })
        json_data['children'].append({
            "type": "bulleted_list_item",
            "bulleted_list_item": {
                "rich_text": [{
                    "type": "text",
                    "text": {
                        "content": ""
                    }
                }]
            }
        })
        json_data['children'].append({       
            "type": "paragraph",
            "paragraph": {
                "rich_text": [{
                    "type": "text",
                    "text": {
                        "content": ""
                    }
                }]
            }   
        })
        json_data['children'].append({
            "type": "divider",
            "divider": {}
        })
    else:
        json_data = {
            'children': []
        }
        createUrl = f'https://api.notion.com/v1/blocks/{blockID}/children'

        json_data['children'].append({
            "type": "heading_1",
            "heading_1": {
                "rich_text": [{
                    "type": "text",
                    "text": {
                        "content": f'Justin'
                    }
                }]
            },
        })
        json_data['children'].append({
            "type": "heading_2",
            "heading_2": {
                "rich_text": [{
                    "type": "text",
                    "text": {
                        "content": "Work Done"
                    }
                }],
                "color": "gray_background"
            }
        })
        json_data['children'].append({
            "type": "paragraph",
            "paragraph": {
                "rich_text": [{
                    "type": "text",
                    "text": {
                        "content": "📝 What tasks did you get done during the last sprint? (Paste links to Issues on Jira)"
                    },
                    "annotations": {
                        "color": "gray",
                    }
                }]
            }
        })

        json_data['children'].append({
            "type": "heading_2",
            "heading_2": {
                "rich_text": [{
                    "type": "text",
                    "text": {
                        "content": "Retrospective"
                    }
                }],
                "color": "gray_background"
            }
        })
        json_data['children'].append({
            "type": "paragraph",
            "paragraph": {
                "rich_text": [{
                    "type": "text",
                    "text": {
                        "content": "🔵 What are some things you liked about or think you did well during the last sprint?"
                    },
                    "annotations": {
                        "color": "gray",
                    }
                }]
            }
        })
        json_data['children'].append({
            "type": "bulleted_list_item",
            "bulleted_list_item": {
                "rich_text": [{
                    "type": "text",
                    "text": {
                        "content": ""
                    }
                }]
            }
        })
        json_data['children'].append({
            "type": "paragraph",
            "paragraph": {
                "rich_text": [{
                    "type": "text",
                    "text": {
                        "content": "🔴 What are some things you felt lacking or think could have gone better?"
                    },
                    "annotations": {
                        "color": "gray",
                    }
                }]
            }
        })
        json_data['children'].append({
            "type": "bulleted_list_item",
            "bulleted_list_item": {
                "rich_text": [{
                    "type": "text",
                    "text": {
                        "content": ""
                    }
                }]
            }
        })
        json_data['children'].append({       
            "type": "paragraph",
            "paragraph": {
                "rich_text": [{
                    "type": "text",
                    "text": {
                        "content": ""
                    }
                }]
            }   
        })

    data = json.dumps(json_data)

    res = requests.request("PATCH", createUrl, headers=headers, data=data)
    print(res.status_code)
    print(res.text)

pageID = createPage(headers, "mainPage.json")
blockID = getTargetBlock(pageID)

member_index = 0
for member in members_notion:
    createPageContent(pageID, member, member_index, False)
    member_index += 1

    if member_index == len(members_notion):
        createPageContent(pageID, "마지막 이름", 0, True)