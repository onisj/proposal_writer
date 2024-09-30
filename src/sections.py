import os
import json
import openai
from .utils import read_word
from .constants import GPTModelLight, GPTModel  



# Define these variables in sections.py
CL_samples = read_word(os.path.join('files', 'Cover letters samples.docx'))
expertise = read_word(os.path.join('files', 'expertise.docx'))
team_qual = read_word(os.path.join('files', 'team_qualification.docx'))
project_und = read_word(os.path.join('files', 'project_understanding.docx'))
technical_app = read_word(os.path.join('files', 'technical_approach.docx'))
workplan = read_word(os.path.join('files', 'work_plan.docx'))
cost_det = read_word(os.path.join('files', 'cost_details.docx'))


def cover_letter(dets, scope):
     # Define the context for the summary
    context = (f'''make a relevant cover letter inspired by the one I did on my previous proposal
[CL_samples] 
and adapt it to this RFP. Here is the information for the purchasing manager, the person in charge and his information : {dets}
and here the scope of service of the RFP :
{scope}
It needs to end with these sentences : 
For any inquiries or additional information, please feel free to contact me directly at (518) 330-8526 or via email at jeff.abraham@grantcityllc.com We are excited about the prospect of becoming your trusted partner and advocate in this crucial endeavor.
Sincerely,
''')

    # Call the OpenAI API to generate a summary
    response =  openai.chat.completions.create(
        model=GPTModelLight,
        temperature=0.0,
        max_tokens=10000,
        messages=[
            {"role": "system", "content": context},
            {"role": "user", "content": CL_samples}
        ]
    )
    
    return response.choices[0].message.content.strip()



def FIRM_QUALIFICATIONS( scope):
     # Define the context for the summary
    context = (f'''Here is the section : FIRM QUALIFICATIONS, EXPERTISE, AND EXPERIENCE from a past proposal : 
               [expertise.docx] Adapt to this new RFP, here is the scope of service of the RFP : {scope}''')

    # Call the OpenAI API to generate a summary
    response =  openai.chat.completions.create(
        model=GPTModelLight,
        temperature=0.0,
        max_tokens=10000,
        messages=[
            {"role": "system", "content": context},
            {"role": "user", "content": expertise}
        ]
    )
    
    return response.choices[0].message.content.strip()



def TEAM_QUALIFICATIONS(scope):
     # Define the context for the summary
    context = (f'''Here is the section “TEAM QUALIFICATIONS, EXPERTISE, AND EXPERIENCE” from a past proposal : [team_qual] 
               It needs to be assertive and persuasive, keep the same tone. Adapt to this new RFP , here is the scope of service of the RFP : {scope}
               ''')

    # Call the OpenAI API to generate a summary
    response =  openai.chat.completions.create(
        model=GPTModelLight,
        temperature=0.0,
        max_tokens=10000,
        messages=[
            {"role": "system", "content": context},
            {"role": "user", "content": team_qual}
        ]
    )
    
    return response.choices[0].message.content.strip()



def Project_understanding(scope, article):
     # Define the context for the summary
    context = (f'''write me a project understanding section for RFP, I am writing a proposal with a section project understanding , 
               here is this section from a past proposal : [project_und] For this new RFP, here is the scope of service / work : {scope} 
               Very important : our company named grant city LLC, adapt the text mentioning that we know all about the disaster event 
               and is ready willing and able to assist (as needed). We add a recent press article about the disaster event, keep the same 
               tone we used in the past proposal , project understanding section. Here is the article : {article}''')

    # Call the OpenAI API to generate a summary
    response =  openai.chat.completions.create(
        model=GPTModelLight,
        temperature=0.0,
        max_tokens=10000,
        messages=[
            {"role": "system", "content": context},
            {"role": "user", "content": project_und}
        ]
    )
    
    return response.choices[0].message.content.strip()



def Technical_approach( scope):
     # Define the context for the summary
    context = (f'''Create a technical approach section for responding to an RFP in a JSON format.
        Detailed methodology to achieve the project's goals, identification of deliverables, key decision points, and any proposed changes to the scope of services.
        If you output a table, output table markdown formatted, any row in its own line. 
        Here is the Technical Approach from a past proposal : 
        [technical_app] 
        Adapt to this new RFP, Approach why our company : grant city LLC about managing the project is the best choice, it needs to be assertive and persuasive . Here is the scope of service of the RFP : 
        [{scope}].
        
        Return Output in JSON format:
        \u007b
            technical_approach : \u007b
                "intro": "make few paragraphs and points for intro taking reference example from [technical_app] provided to you.",
                "table": [
                    "headers": ["Phase", "Objective", "Deliverables", "Key Decision Points"],
                    "rows": [
                        ["Phase 1", "Objective", "Deliverables", "Key Decision Points"],
                        ["Phase 2", "Objective", "Deliverables", "Key Decision Points"],
                            .
                            .
                        //...more rows as needed
                    ]
                ],
                "conclusion": "make few paragraphs for conclusion taking reference example from [technical_app] provided to you."
            \u007b
        \u007b
    ''')

    # Call the OpenAI API to generate a summary
    response =  openai.chat.completions.create(
        model=GPTModelLight,
        temperature=0.0,
        max_tokens=10000,
        messages=[
            {"role": "system", "content": context},
            {"role": "user", "content": technical_app}
        ]
    )
    
    response_content = response.choices[0].message.content
    cleaned_response = response_content.replace('```json\n', '').replace('\n```', '')

    # Parse the JSON response
    try:
        structured_data = json.loads(cleaned_response)
    except json.JSONDecodeError as e:
        print("Failed to decode JSON:", e)
        return None
    return structured_data
 


def Work_plan(scope): 
    
     #extracting existing teams for proposed title and role.
    path = os.path.join('files', 'work_plan.docx')
    doc = Document(path)
    staff_titles_dict = {}
    
    for table in doc.tables:
        for row in table.rows[1:]:
            name = row.cells[0].text.strip()
            title = row.cells[1].text.strip()
            staff_titles_dict[name] = title
    
    names_and_titles = [f"{name}: {title}" for name, title in staff_titles_dict.items()]
    
     # Define the context for the summary 
    context = (f'''Create a work plan section for responding to an RFP in JSON format. 
    The work plan should include a detailed project schedule, a matrix of personnel tasks, and estimated effort in hours. 
    **Use the existing team members (names, titles, and roles) from the previous project as provided here {names_and_titles}. 
    Do not create dummy names or new data for the team. Only adapt their specific tasks, duties, and hours to reflect the new RFP's scope of services.**

    Here is the work plan from a past proposal:
        [workplan]

    CRITICAL:**Ensure that the team members' names, titles, and roles remain unchanged. However, the specific duties, tasks, and hours allocated may change 
    according to the new RFP requirements.**

    Here is the scope of service of the RFP: 
        [{scope}]

    Remember: The output should be similar in length and content to the past proposal, but adapt to the new RFP's requirements, maintaining a similar matrix structure.

    Output Format:
    \u007b
        work_plan : \u007b
            "intro": "Take reference example from [workplan] in length and tone provided to you for intro.",
            "table": [
                "headers": ["Proposed Staff", "Proposed Title", "Proposed Duty Statement", "SPECIFIC SOW ITEM PERFORMED BY RELEVANT STAFF(AS REQUESTED IN THE RFP)",
                "ANNUAL ALLOCATION OF HOURS PER TASK"],
                "rows": [
                    ["Name of existing staff", "Title", "Duty", "Sow item by relevant staff","Annual Hours allocation"],
                    ["Name of existing staff", "Title", "Duty", "Sow item by relevant staff","Annual Hours allocation"],
                        .
                        .
                    ["Name of existing staff", "Title", "Duty", "Sow item by relevant staff","Annual Hours allocation"],
                ]
            ],
            "conclusion": "Take reference example from [workplan] in length and tone provided to you for conclusion."
        \u007b
    \u007b
    ''')

    # Call the OpenAI API to generate a summary
    response =  openai.chat.completions.create(
        model=GPTModelLight,
        temperature=0.0,
        max_tokens=10000,
        messages=[
            {"role": "system", "content": context},
            {"role": "user", "content": workplan}
        ]
    )
    
    response_content = response.choices[0].message.content
    cleaned_response = response_content.replace('```json\n', '').replace('\n```', '')

    # Parse the JSON response
    try:
        structured_data = json.loads(cleaned_response)
    except json.JSONDecodeError as e:
        print("Failed to decode JSON:", e)
        return None
    return structured_data


def COST_PROPOSAL(scope, budget):
     #extracting existing teams for proposed title and role.
    path = os.path.join('files', 'cost_details.docx')
    doc = Document(path)
    staff_titles_dict = {}
    
    for table in doc.tables:
        for row in table.rows[1:]:
            name = row.cells[0].text.strip()
            title = row.cells[1].text.strip()
            staff_titles_dict[name] = title
    
    names_and_titles = [f"{name}: {title}" for name, title in staff_titles_dict.items()]
         
     # Define the context for the summary
    context = (f'''Create a cost proposal section for responding to an RFP in JSON format. 
               
        CRITICAL:**Use the same team data/existing (names, titles, and roles) as provided here {names_and_titles} without generating new names or changing the team structure provided to you in [cost_det] ""
        Adapt the cost details, including rates, hours, and other expenses, to fit the new yearly budget and scope of services.

        Here is the cost proposal from a previous proposal we answered:
        [cost_det]

        Ensure the cost proposal reflects the existing team, but adjust their specific cost details to the new budget:
        {budget}.

        Here is the scope of service of the RFP:
        [{scope}]
        Output Format:
        \u007b
            "cost_proposal" : \u007b
                "intro": "Take reference example from [cost_det] in length and tone provided to you for intro.",
                "table1": [
                    "headers": ["Proposed Personnel", "Position Title", "Direct Labor Rate (Based on a normal 8-hour, 40 hours per week schedule)", "Indirect Labor Costs ","Professional Fee (Profit)","Other Direct Costs (ODC)","Fully Burdened Hourly Rate"],
                    "rows": [
                        ["Name of personnel", "Position", "Rate", "Indirect Labour cost","Profit","ODC","Hourly Rate"],
                        ["Name of personnel", "Position", "Rate", "Indirect Labour cost","Profit","ODC","Hourly Rate"],
                            .
                            .
                        ["Name of personnel", "Position", "Rate", "Indirect Labour cost","Profit","ODC","Hourly Rate"],
                    ]
                "table2": [
                    "headers": ["Initial Contract Period","Total Annual Not to Exceed Amount"],
                    "rows": [
                        ["Year", "Amount"]
                    ]
                ],
                "conclusion": "Take reference example from [cost_det] and provide exception and assumption for conclusion."
            \u007b
        \u007b   
    ''')

    # Call the OpenAI API to generate a summary
    response =  openai.chat.completions.create(
        model=GPTModelLight,
        temperature=0.0,
        max_tokens=10000,
        messages=[
            {"role": "system", "content": context},
            {"role": "user", "content": cost_det}
        ]
    )
    
    response_content = response.choices[0].message.content
    cleaned_response = response_content.replace('```json\n', '').replace('\n```', '')

    # Parse the JSON response
    try:
        structured_data = json.loads(cleaned_response)
    except json.JSONDecodeError as e:
        print("Failed to decode JSON:", e)
        return None
    return structured_data



def ref(scope):
     # Define the context for the summary
    context = (f'''I am creating a proposal to answer an RFP, my company GCC (Grant City LLC), we are writing the client reference section. I will send you an example of a past answer from a past proposal and you will adapt to the new RFP, with the new client name etc…
Here is a text from a past proposal for a previous client : 
GCC has provided descriptions and client references for projects of similar scope and size to the City’s project performed within the past three years. The City of Santy Cruz may contact any one of our client references during the evaluation phase. GCC manages more than 4,000 projects each year and has an extensive list of satisfied clients and award-winning performance on a wide range of diverse engineering, program management, and construction management assignments, including disaster recovery projects across the County. Figure 6 and Figure 7 provide a summary of our references and their similarity in scope, size, and complexity to Santa Cruz.
And here is the scope of service of the RFP : 
{scope}.

''')

    # Call the OpenAI API to generate a summary
    response =  openai.chat.completions.create(
        model=GPTModelLight,
        temperature=0.0,
        max_tokens=10000,
        messages=[
            {"role": "system", "content": context},
        ]
    )
    
    return response.choices[0].message.content.strip()
