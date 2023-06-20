import subprocess
import base64
import os
import css_inline

file_path = os.path.dirname(__file__)

languages = ["markup", "css", "clike", "javascript", "abap", "abnf", "actionscript", "ada", "apacheconf", "apl", "applescript", "arduino", "arff", "asciidoc", "asm6502", "aspnet", "autohotkey", "autoit", "bash", "basic", "batch", "bison", "bnf", "brainfuck", "bro", "c", "csharp", "cpp", "cil", "coffeescript", "cmake", "clojure", "crystal", "csp", "css-extras", "d", "dart", "diff", "django", "docker", "ebnf", "eiffel", "ejs", "elixir", "elm", "erb", "erlang", "fsharp", "flow", "fortran", "gcode", "gedcom", "gherkin", "git", "glsl", "gml", "go", "graphql", "groovy", "haml", "handlebars", "haskell", "haxe", "hcl", "http", "hpkp", "hsts", "ichigojam", "icon", "inform7", "ini", "io", "j", "java", "javadoc", "javadoclike", "javastacktrace", "jolie", "jsdoc", "js-extras", "json", "jsonp", "json5", "julia", "keyman", "kotlin", "latex", "less", "liquid", "lisp", "livescript", "lolcode", "lua", "makefile", "markdown", "markup-templating", "matlab", "mel", "mizar", "monkey", "n1ql", "n4js", "nand2tetris-hdl", "nasm", "nginx", "nim", "nix", "nsis", "objectivec", "ocaml", "opencl", "oz", "parigp", "parser", "pascal", "perl", "php", "phpdoc", "php-extras", "plsql", "powershell", "processing", "prolog", "properties", "protobuf", "pug", "puppet", "pure", "python", "q", "qore", "r", "jsx", "tsx", "renpy", "reason", "regex", "rest", "rip", "roboconf", "ruby", "rust", "sas", "sass", "scss", "scala", "scheme", "smalltalk", "smarty", "sql", "soy", "stylus", "swift", "tap", "tcl", "textile", "toml", "tt2", "twig", "typescript", "t4-cs", "t4-vb", "t4-templating", "vala", "vbnet", "velocity", "verilog", "vhdl", "vim", "visual-basic", "wasm", "wiki", "xeora", "xojo", "xquery", "yaml", ]

def highlight(code, language):
    code_b64 = base64.b64encode(code.encode()).decode('utf-8')
    
    node_command = ['node', os.path.join(file_path,'highlight.js'), code_b64, language]

    result = subprocess.run(node_command, capture_output=True, text=True)

    output = result.stdout.strip()
    if (result.stderr.strip() != ''):
        raise Exception(result.stderr.strip())

    code = base64.b64decode(output).decode('utf-8')
    css = open(os.path.join(file_path, 'prism.css')).read()
    inline = css_inline.inline(f'<style>{css}</style>'+code)
    code = inline[inline.find('<body>')+len('<body>'):][:-len('</body></html>')]
    return code


if __name__ == '__main__':
    code = "const message = 'Hello, World!';"
    code_hl = (highlight(code, 'javascript'))
    print(code_hl)
