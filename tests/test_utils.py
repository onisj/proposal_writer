from src.utils import cleaner, extract_text_from_pdf, read_word

def test_cleaner():
    assert cleaner("**Hello# World**") == "Hello World"

def test_extract_text_from_pdf():
    # Use a sample PDF file for testing
    with open("tests/sample.pdf", "rb") as f:
        pdf_text = extract_text_from_pdf(f)
    assert "sample text" in pdf_text.lower() # Check for expected content

def test_read_word():
    # Use a sample DOCX file for testing
    word_text = read_word("tests/sample.docx")
    assert "sample text" in word_text.lower() # Check for expected content
