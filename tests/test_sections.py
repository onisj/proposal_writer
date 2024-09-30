from src.sections import cover_letter, FIRM_QUALIFICATIONS, TEAM_QUALIFICATIONS, Project_understanding

def test_cover_letter():
    # Provide sample inputs for testing
    dets = "Sample purchasing manager details"
    scope = "Sample scope of work"
    cover_letter_text = cover_letter(dets, scope)
    assert "grantcityllc.com" in cover_letter_text # Check for expected content

def test_FIRM_QUALIFICATIONS():
    # Provide sample inputs for testing
    scope = "Sample scope of work"
    fq_text = FIRM_QUALIFICATIONS(scope)
    assert "expertise" in fq_text.lower() # Check for expected content

def test_TEAM_QUALIFICATIONS():
    # Provide sample inputs for testing
    scope = "Sample scope of work"
    tq_text = TEAM_QUALIFICATIONS(scope)
    # Check if the generated text contains keywords related to team qualifications
    assert any(keyword in tq_text.lower() for keyword in ["team", "qualifications", "experience"]) 

def test_Project_understanding():
    # Provide sample inputs for testing
    scope = "Sample scope of work"
    article = "Sample article about the disaster"
    pu_text = Project_understanding(scope, article)
    # Check if the generated text contains keywords related to project understanding
    assert any(keyword in pu_text.lower() for keyword in ["project", "understanding", "disaster"]) 
