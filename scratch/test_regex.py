import re

def test_split(text):
    # Regex that matches the label part: 
    # optional prefix (a-e followed by separator), 
    # followed by "test case", 
    # followed by the keyword, 
    # followed by optional separators like space, dot, dash, colon.
    RE_LABEL = re.compile(r'^(([a-e][\s\.\)]\s*)?(test\s*case\s*)?(name|description|execution|observation|evidence)[\w\s]*?[\s\.\-\:]+)', re.IGNORECASE)
    
    m = RE_LABEL.match(text)
    if m:
        label = m.group(1)
        content = text[m.end():].strip()
        print(f"TEXT: {text}")
        print(f"LABEL: '{label}'")
        print(f"CONTENT: '{content}'")
        print("-" * 20)
    else:
        print(f"TEXT: {text} -> NO MATCH")

test_split("b. Test Case description: To ensure that only...")
test_split("b. Test Case description To ensure that only...")
test_split("Test Case description: To ensure that only...")
test_split("Test Case description To ensure that only...")
test_split("a. Test Case Name: My Test")
test_split("a. Test Case Name My Test")
test_split("c. Execution: Step 1")
test_split("c. Execution Step 1")
