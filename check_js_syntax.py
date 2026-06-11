def check_braces(filename):
    with open(filename, 'r', encoding='utf-8') as f:
        content = f.read()
    
    stack = []
    mapping = {')': '(', '}': '{', ']': '['}
    lines = content.split('\n')
    
    # We want to keep track of line and char numbers
    line_num = 1
    char_num = 1
    
    in_string = False
    string_char = None
    in_comment = False
    in_multiline_comment = False
    
    i = 0
    while i < len(content):
        c = content[i]
        
        # Track line numbers
        if c == '\n':
            line_num += 1
            char_num = 1
            i += 1
            if in_comment:
                in_comment = False
            continue
        
        # Handle multiline comments
        if in_multiline_comment:
            if i + 1 < len(content) and content[i] == '*' and content[i+1] == '/':
                in_multiline_comment = False
                i += 2
                char_num += 2
            else:
                i += 1
                char_num += 1
            continue
            
        # Handle single line comments
        if in_comment:
            i += 1
            char_num += 1
            continue
            
        # Handle strings
        if in_string:
            if c == '\\':
                # Skip next char (escaped)
                i += 2
                char_num += 2
                continue
            if c == string_char:
                in_string = False
            i += 1
            char_num += 1
            continue
            
        # Comment markers
        if i + 1 < len(content) and content[i] == '/' and content[i+1] == '/':
            in_comment = True
            i += 2
            char_num += 2
            continue
        if i + 1 < len(content) and content[i] == '/' and content[i+1] == '*':
            in_multiline_comment = True
            i += 2
            char_num += 2
            continue
            
        # String markers
        if c in ['"', "'", '`']:
            in_string = True
            string_char = c
            i += 1
            char_num += 1
            continue
            
        # Braces
        if c in ['(', '{', '[']:
            stack.append((c, line_num, char_num))
        elif c in [')', '}', ']']:
            if not stack:
                print(f"Unmatched closing '{c}' at line {line_num}, column {char_num}")
                return False
            top_char, top_line, top_col = stack.pop()
            if mapping[c] != top_char:
                print(f"Mismatch: '{c}' at line {line_num}, col {char_num} does not match '{top_char}' from line {top_line}, col {top_col}")
                return False
                
        i += 1
        char_num += 1
        
    if stack:
        print(f"Unmatched opening brackets remaining: {len(stack)}")
        for char, line, col in stack[-5:]:
            print(f"  - '{char}' at line {line}, col {col}")
        return False
        
    print("All brackets, braces, and parentheses are balanced successfully!")
    return True

check_braces(r"C:\Users\hp\OneDrive\Desktop\fronted\app.js")
