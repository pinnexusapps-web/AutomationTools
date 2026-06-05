import os
import re
import html
import sys

# ----------------------------------------------------------------------
# VBA keyword set (case‑insensitive)
# ----------------------------------------------------------------------
VBA_KEYWORDS = {
    'sub', 'end', 'function', 'property', 'get', 'let', 'set', 'dim', 'as',
    'public', 'private', 'friend', 'static', 'const', 'byval', 'byref',
    'optional', 'paramarray', 'new', 'nothing', 'is', 'like', 'mod', 'and',
    'or', 'xor', 'not', 'if', 'then', 'else', 'elseif', 'for', 'to', 'step',
    'next', 'do', 'loop', 'while', 'wend', 'select', 'case', 'with', 'exit',
    'goto', 'resume', 'on', 'error', 'msgbox', 'format', 'vlookup', 'split',
    'cells', 'activeworkbook', 'sheets', 'range', 'rows', 'count', 'end',
    'up', 'xlup', 'xlvalues', 'xlwhole', 'address', 'value', 'vbcritical',
    'vbinformation', 'vbcrlf', 'worksheet', 'nothing', 'boolean',
    'integer', 'long', 'double', 'string', 'variant', 'object', 'byte',
    'single', 'currency', 'decimal', 'date', 'time'
}

# ----------------------------------------------------------------------
# Token specification (order matters)
# ----------------------------------------------------------------------
TOKEN_PATTERNS = [
    ('COMMENT',   r"'.*?$"),
    ('STRING',    r'"(?:[^"]|"")*"'),
    ('NUMBER',    r'\b\d+(?:\.\d+)?\b'),
    ('IDENTIFIER',r'[a-zA-Z_][a-zA-Z0-9_]*'),
    ('OPERATOR',  r'<>|<=|>=|=|>|<|[+\-*/&,;:()\[\]{}]'),
    ('WHITESPACE',r'\s+'),
]

TOKEN_RE = re.compile('|'.join(f'(?P<{name}>{pattern})' for name, pattern in TOKEN_PATTERNS), re.MULTILINE)

# ----------------------------------------------------------------------
def tokenize(vba_code: str):
    for mo in TOKEN_RE.finditer(vba_code):
        for name, _ in TOKEN_PATTERNS:
            token = mo.group(name)
            if token is not None:
                yield name, token
                break

# ----------------------------------------------------------------------
def vba_to_html(vba_code: str) -> str:
    tokens = list(tokenize(vba_code))
    output = []
    expect_fn = False
    i = 0
    n = len(tokens)

    while i < n:
        ttype, tval = tokens[i]

        if ttype == 'WHITESPACE':
            output.append(tval)
            i += 1
            continue
        if ttype == 'COMMENT':
            output.append(f'<span class="cm">{html.escape(tval)}</span>')
            i += 1
            continue
        if ttype == 'STRING':
            output.append(f'<span class="str">{html.escape(tval)}</span>')
            i += 1
            continue
        if ttype == 'NUMBER':
            output.append(f'<span class="num">{html.escape(tval)}</span>')
            i += 1
            continue
        if ttype == 'OPERATOR':
            output.append(html.escape(tval))
            i += 1
            continue
        if ttype == 'IDENTIFIER':
            lower_val = tval.lower()

            # label detection: identifier followed by colon (skip whitespace)
            is_label = False
            j = i + 1
            while j < n and tokens[j][0] == 'WHITESPACE':
                j += 1
            if j < n and tokens[j][0] == 'OPERATOR' and tokens[j][1] == ':':
                is_label = True

            if expect_fn and not is_label:
                output.append(f'<span class="fn">{html.escape(tval)}</span>')
                expect_fn = False
                i += 1
                continue

            if is_label:
                # consume the colon as well
                colon_token = tokens[j]
                output.append(f'<span class="fn">{html.escape(tval)}{html.escape(colon_token[1])}</span>')
                i = j + 1
                continue

            if lower_val in VBA_KEYWORDS:
                output.append(f'<span class="kw">{html.escape(tval)}</span>')
                if lower_val in ('sub', 'function'):
                    expect_fn = True
                i += 1
                continue

            # default: variable / user identifier
            output.append(f'<span class="var">{html.escape(tval)}</span>')
            i += 1
            continue

        # fallback (should not occur)
        output.append(html.escape(tval))
        i += 1

    return ''.join(output)

# ----------------------------------------------------------------------
def convert_vba_file_to_html(input_path: str, output_path: str = None):
    """
    Reads a VBA text file and writes a syntax‑highlighted HTML file.
    If output_path is None, the HTML file is saved in the same folder
    as the input file, with the same base name but .html extension.
    """
    if not os.path.isfile(input_path):
        raise FileNotFoundError(f"Input file not found: {input_path}")

    # Determine output path if not given
    if output_path is None:
        base, _ = os.path.splitext(input_path)
        output_path = base + ".html"

    # Read entire VBA source (supports large files, no truncation)
    with open(input_path, 'r', encoding='utf-8') as f:
        vba_code = f.read()

    # Convert to highlighted HTML fragment
    highlighted_fragment = vba_to_html(vba_code)

    # Build complete HTML document
    title = os.path.basename(input_path)
    css = """
    <style>
        .kw { color: #0000FF; font-weight: bold; }
        .fn { color: #008080; font-weight: bold; }
        .var { color: #2E8B57; }
        .str { color: #CC0000; }
        .num { color: #FF00FF; }
        .cm { color: #008000; font-style: italic; }
        pre {
            background-color: #f4f4f4;
            border: 1px solid #ddd;
            padding: 10px;
            font-family: 'Courier New', monospace;
            font-size: 14px;
            line-height: 1.4;
            overflow-x: auto;
        }
        code { font-family: inherit; }
    </style>
    """

    full_html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>{html.escape(title)}</title>
{css}
</head>
<body>
<pre><code>{highlighted_fragment}</code></pre>
</body>
</html>"""

    # Write HTML file
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(full_html)

    print(f"Successfully converted: {input_path} -> {output_path}")

# ----------------------------------------------------------------------
if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python vba_to_html.py <path_to_vba_text_file>")
        sys.exit(1)

    input_file = sys.argv[1]
    try:
        convert_vba_file_to_html(input_file)
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)