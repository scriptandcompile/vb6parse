#!/usr/bin/env python3
"""
Generate VB6 Library Reference documentation from Rust source files.

This script extracts module documentation (//! comments) from Rust files in
src/syntax/library/ and generates HTML pages for the GitHub Pages documentation site.
"""

import argparse
import json
import os
import re
import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass, field

try:
    import markdown
    from markdown.extensions import fenced_code, tables, toc, nl2br
except ImportError:
    print("Error: markdown library not found. Install with: pip install markdown")
    sys.exit(1)


@dataclass
class LibraryItem:
    """Represents a single VB6 library function or statement."""
    name: str
    category: str
    subcategory: Optional[str]
    file_path: Path
    doc_content: str
    item_type: str  # 'function' or 'statement'
    
    @property
    def slug(self) -> str:
        """Generate URL-friendly slug from name."""
        return self.name.lower().replace('$', '_dollar').replace('#', '_hash')
    
    @property
    def html_filename(self) -> str:
        """Generate HTML filename."""
        return f"{self.slug}.html"


@dataclass
class Category:
    """Represents a category of library items."""
    name: str
    display_name: str
    description: str
    items: List[LibraryItem] = field(default_factory=list)
    
    @property
    def slug(self) -> str:
        """Generate URL-friendly slug."""
        return self.name


def extract_module_docs(file_path: Path) -> Optional[str]:
    """
    Extract module documentation (//! comments) from a Rust file.
    
    Args:
        file_path: Path to the .rs file
        
    Returns:
        Extracted documentation as markdown string, or None if no docs found
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
    except Exception as e:
        print(f"Warning: Could not read {file_path}: {e}")
        return None
    
    doc_lines = []
    in_module_doc = False
    
    for line in lines:
        stripped = line.strip()
        
        # Check for module doc comment
        if stripped.startswith('//! '):
            in_module_doc = True
            # Remove //!  prefix, but not leading whitespace which might be important in code blocks.
            content = stripped[4:]
            doc_lines.append(content)
        elif in_module_doc:
            # Stop at first non-doc line after docs started
            if stripped and not stripped.startswith('//'):
                break
            # Continue if empty line or regular comment
            elif not stripped:
                doc_lines.append('')
    
    if doc_lines:
        # Join and clean up
        content = '\n'.join(doc_lines).strip()
        return content
    
    return None


def parse_library_structure(src_dir: Path) -> Tuple[List[Category], List[Category]]:
    """
    Parse the library directory structure and extract all documentation.
    
    Args:
        src_dir: Path to src/syntax/library/
        
    Returns:
        Tuple of (function_categories, statement_categories)
    """
    functions_dir = src_dir / "functions"
    statements_dir = src_dir / "statements"
    
    function_categories = parse_category_dir(functions_dir, "function")
    statement_categories = parse_category_dir(statements_dir, "statement")
    
    return function_categories, statement_categories


def parse_category_dir(category_dir: Path, item_type: str) -> List[Category]:
    """
    Parse a category directory (functions or statements) and extract items.
    
    Args:
        category_dir: Path to category directory
        item_type: 'function' or 'statement'
        
    Returns:
        List of Category objects
    """
    if not category_dir.exists():
        print(f"Warning: Directory not found: {category_dir}")
        return []
    
    categories = []
    
    # Get category metadata from parent mod.rs
    category_descriptions = get_category_descriptions(category_dir / "mod.rs")
    
    # Iterate through subdirectories (categories)
    for subdir in sorted(category_dir.iterdir()):
        if not subdir.is_dir():
            continue
        
        category_name = subdir.name
        display_name = category_name.replace('_', ' ').title()
        description = category_descriptions.get(category_name, f"{display_name} {item_type}s")
        
        category = Category(
            name=category_name,
            display_name=display_name,
            description=description
        )
        
        # Parse all .rs files in this category (except mod.rs)
        for rs_file in sorted(subdir.glob("*.rs")):
            if rs_file.name == "mod.rs":
                continue
            
            item_name = rs_file.stem
            doc_content = extract_module_docs(rs_file)
            
            if doc_content:
                item = LibraryItem(
                    name=item_name,
                    category=category_name,
                    subcategory=None,
                    file_path=rs_file,
                    doc_content=doc_content,
                    item_type=item_type
                )
                category.items.append(item)
        
        if category.items:
            categories.append(category)
    
    return categories


def get_category_descriptions(mod_file: Path) -> Dict[str, str]:
    """
    Extract category descriptions from mod.rs comments.
    
    Args:
        mod_file: Path to mod.rs file
        
    Returns:
        Dictionary mapping category names to descriptions
    """
    descriptions = {}
    
    if not mod_file.exists():
        return descriptions
    
    try:
        with open(mod_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Look for patterns like: - **`category`** - Description
        pattern = r'-\s+\*\*`?(\w+)`?\*\*\s*[-:]?\s*(.+?)(?=\n|$)'
        for match in re.finditer(pattern, content):
            category = match.group(1)
            description = match.group(2).strip()
            descriptions[category] = description
    
    except Exception as e:
        print(f"Warning: Could not parse {mod_file}: {e}")
    
    return descriptions


def markdown_to_html(md_content: str) -> str:
    """
    Convert markdown content to HTML with syntax highlighting.
    
    Maps VB6/VB code blocks to vbnet for highlight.js compatibility.
    
    Args:
        md_content: Markdown content
        
    Returns:
        HTML string
    """
    import html as html_module

    # --- Code block extraction ---
    code_blocks = []
    def save_code_block(match):
        lang = match.group(1).strip() if match.group(1) else ''
        code = match.group(2)
        if lang in ('vb', 'vb6'):
            lang = 'vbnet'
        escaped_code = html_module.escape(code)
        if lang:
            html_code = f'<pre><code class="language-{lang}">{escaped_code}</code></pre>'
        else:
            html_code = f'<pre><code>{escaped_code}</code></pre>'
        placeholder = f'\n\nXOXOCODEBLOCKXOXO{len(code_blocks)}XOXOENDXOXO\n\n'
        code_blocks.append(html_code)
        return placeholder
    md_content = re.sub(r'```(\w*)\n(.*?)\n```', save_code_block, md_content, flags=re.DOTALL)

    # --- Table extraction ---
    # Markdown tables: lines with | and at least one header separator (| ---)
    # We'll extract blocks of lines that look like tables, separated by empty lines
    table_blocks = []
    table_pattern = re.compile(r'((?:^[|].*\n)+^[| ]*[-:]+[-| :]*\n(?:^[|].*\n)+)', re.MULTILINE)
    def save_table_block(match):
        table_md = match.group(1)
        placeholder = f'\n\nXOXOTABLEBLOCKXOXO{len(table_blocks)}XOXOENDXOXO\n\n'
        table_blocks.append(table_md)
        return placeholder
    md_content = table_pattern.sub(save_table_block, md_content)

    # Convert remaining markdown
    md = markdown.Markdown(
        extensions=[
            'tables',
            'toc'
        ]
    )
    html_output = md.convert(md_content)

    # Restore table blocks (convert to HTML using markdown, then insert)
    for i, table_md in enumerate(table_blocks):
        placeholder = f'XOXOTABLEBLOCKXOXO{i}XOXOENDXOXO'
        # Convert table markdown to HTML only (no toc)
        table_html = markdown.markdown(table_md, extensions=['tables'])
        # Remove wrapping <p> tags if present
        html_output = html_output.replace(f'<p>{placeholder}</p>', table_html)
        html_output = html_output.replace(placeholder, table_html)

    # Restore code blocks
    for i, code_block in enumerate(code_blocks):
        placeholder = f'XOXOCODEBLOCKXOXO{i}XOXOENDXOXO'
        html_output = html_output.replace(f'<p>{placeholder}</p>', code_block)
        html_output = html_output.replace(placeholder, code_block)

    return html_output


def generate_html_page(title: str, content: str, breadcrumbs: List[Tuple[str, str]], 
                      output_path: Path, base_path: str = "../../") -> None:
    """
    Generate an HTML page with consistent styling.
    
    Args:
        title: Page title
        content: HTML content for main section
        breadcrumbs: List of (text, url) tuples for breadcrumb navigation
        output_path: Path where HTML file will be saved
        base_path: Relative path to docs root (for CSS/JS)
    """
    breadcrumb_html = ' / '.join([
        f'<a href="{base_path}{url}">{text}</a>' if url else text
        for text, url in breadcrumbs
    ])
    
    html_template = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="description" content="VB6Parse Library Reference - {title}">
    <title>{title} - VB6Parse Library Reference</title>
    <link rel="stylesheet" href="{base_path}assets/css/style.css">
    <link rel="stylesheet" href="{base_path}assets/css/docs-style.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.9.0/styles/github-dark.min.css">
    <script src="{base_path}assets/js/theme-switcher.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.9.0/highlight.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.9.0/languages/vbnet.min.js"></script>
    <script>hljs.highlightAll();</script>
</head>
<body>
    <header class="docs-header">
        <div class="container">
            <h1>{breadcrumb_html}</h1>
            <p class="tagline">VB6 Library Reference</p>
        </div>
    </header>

    <nav class="docs-nav">
        <div class="container">
            <a href="{base_path}index.html">Home</a>
            <a href="{base_path}library/index.html">Library Reference</a>
            <a href="{base_path}documentation.html">Documentation</a>
            <a href="https://docs.rs/vb6parse" target="_blank">API Docs</a>
            <a href="https://github.com/scriptandcompile/vb6parse" target="_blank">GitHub</a>
            <button id="theme-toggle" class="theme-toggle" aria-label="Toggle theme">
                <span class="theme-icon">🌙</span>
            </button>
        </div>
    </nav>

    <main class="container">
        {content}
    </main>

    <footer>
        <div class="container">
            <p>&copy; 2024-2026 VB6Parse Contributors. Licensed under the MIT License.</p>
        </div>
    </footer>
</body>
</html>
"""
    
    # Ensure output directory exists
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    # Write HTML file
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_template)
    
    print(f"Generated: {output_path}")


def generate_library_index(functions: List[Category], statements: List[Category], 
                          output_dir: Path) -> None:
    """Generate main library index page."""
    
    # Count total items
    total_functions = sum(len(cat.items) for cat in functions)
    total_statements = sum(len(cat.items) for cat in statements)
    
    content = f"""
        <section id="library-overview">
            <h2>VB6 Library Reference</h2>
            <p>
                Complete reference documentation for Visual Basic 6 built-in functions and statements.
                This reference covers {total_functions} functions organized in {len(functions)} categories 
                and {total_statements} statements in {len(statements)} categories.
            </p>
            
            <div class="info-box" style="margin: 2rem 0;">
                <strong>📖 About This Reference:</strong> This documentation is automatically generated from 
                the VB6Parse source code. Each entry includes syntax, parameters, return values, remarks, 
                and comprehensive examples to help you understand VB6's built-in library.
            </div>
        </section>

        <section id="functions" style="margin-top: 3rem;">
            <h2>Functions ({total_functions} items)</h2>
            <p>VB6 functions return values and can be used in expressions.</p>
            
            <div class="feature-grid">
"""
    
    for category in functions:
        content += f"""
                <a href="functions/{category.slug}/index.html" class="feature-card" style="text-decoration: none; color: inherit;">
                    <h2>{category.display_name}</h2>
                    <p>{category.description}</p>
                    <small>{len(category.items)} functions</small>
                </a>
"""
    
    content += """
            </div>
        </section>

        <section id="statements" style="margin-top: 3rem;">
            <h2>Statements (""" + str(total_statements) + """ items)</h2>
            <p>VB6 statements perform actions and control program flow.</p>
            
            <div class="feature-grid">
"""
    
    for category in statements:
        content += f"""
                <a href="statements/{category.slug}/index.html" class="feature-card" style="text-decoration: none; color: inherit;">
                    <h3>{category.display_name}</h3>
                    <p>{category.description}</p>
                    <small>{len(category.items)} statements</small>
                </a>
"""
    
    content += """
            </div>
        </section>
"""
    
    breadcrumbs = [
        ("VB6Parse", "index.html"),
        ("Library Reference", "")
    ]
    
    generate_html_page(
        "Library Reference",
        content,
        breadcrumbs,
        output_dir / "index.html",
        base_path="../"
    )


def generate_category_index(category: Category, item_type: str, output_dir: Path) -> None:
    """Generate category index page listing all items in the category."""
    
    content = f"""
        <section id="category-overview" style="margin-bottom: 2rem;">
            <h2>{category.display_name}</h2>
            <p style="font-size: 1.1rem; margin: 1rem 0;">{category.description}</p>
            <p style="color: var(--primary-color); font-weight: 600;">{len(category.items)} {item_type}s in this category</p>
        </section>

        <section id="items-list" style="margin-top: 2rem;">
            <div class="feature-grid">
"""
    
    for item in sorted(category.items, key=lambda x: x.name.lower()):
        # Extract first line or title from documentation
        first_line = item.doc_content.split('\n')[0].strip('#').strip()
        # Limit description length
        if len(first_line) > 100:
            first_line = first_line[:97] + "..."
        
        content += f"""
                <a href="{item.html_filename}" class="feature-card" style="text-decoration: none; color: inherit;">
                    <h3>{item.name}</h3>
                    <p style="margin-bottom: 0; font-size: 0.95rem;">{first_line}</p>
                </a>
"""
    
    content += """
            </div>
        </section>
"""
    
    breadcrumbs = [
        ("VB6Parse", "index.html"),
        ("Library", "library/index.html"),
        (f"{item_type.title()}s", None),
        (category.display_name, "")
    ]
    
    generate_html_page(
        f"{category.display_name} - {item_type.title()}s",
        content,
        breadcrumbs,
        output_dir / "index.html",
        base_path="../../../"
    )


def generate_item_page(item: LibraryItem, category: Category, output_dir: Path) -> None:
    """Generate individual function/statement documentation page."""
    
    # Convert markdown to HTML
    html_content = markdown_to_html(item.doc_content)
    
    content = f"""
        <article class="library-item">
            {html_content}
        </article>
        
        <div style="margin-top: 3rem; padding-top: 2rem; border-top: 1px solid var(--border-color);">
            <p>
                <a href="index.html">← Back to {category.display_name}</a> |
                <a href="../index.html">View all {item.item_type}s</a>
            </p>
        </div>
"""
    
    breadcrumbs = [
        ("VB6Parse", "index.html"),
        ("Library", "library/index.html"),
        (category.display_name, f"library/{item.item_type}s/{category.slug}/index.html"),
        (item.name, "")
    ]
    
    generate_html_page(
        f"{item.name} - {category.display_name}",
        content,
        breadcrumbs,
        output_dir / item.html_filename,
        base_path="../../../"
    )


def generate_search_index(functions: List[Category], statements: List[Category], 
                         output_dir: Path) -> None:
    """Generate search index JSON file for client-side search."""
    
    search_data = []
    
    # Add functions
    for category in functions:
        for item in category.items:
            # Extract plain text from markdown for search
            plain_text = re.sub(r'[#*`\[\]]', '', item.doc_content)
            plain_text = ' '.join(plain_text.split()[:100])  # First 100 words
            
            search_data.append({
                'name': item.name,
                'type': 'function',
                'category': category.display_name,
                'url': f'functions/{category.slug}/{item.html_filename}',
                'description': plain_text[:200]
            })
    
    # Add statements
    for category in statements:
        for item in category.items:
            plain_text = re.sub(r'[#*`\[\]]', '', item.doc_content)
            plain_text = ' '.join(plain_text.split()[:100])
            
            search_data.append({
                'name': item.name,
                'type': 'statement',
                'category': category.display_name,
                'url': f'statements/{category.slug}/{item.html_filename}',
                'description': plain_text[:200]
            })
    
    # Write JSON
    search_index_path = output_dir / "search-index.json"
    with open(search_index_path, 'w', encoding='utf-8') as f:
        json.dump(search_data, f, indent=2)
    
    print(f"Generated search index: {search_index_path}")


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description="Generate VB6 Library Reference documentation from Rust source files"
    )
    parser.add_argument(
        "--src",
        type=Path,
        default=Path("src/syntax/library"),
        help="Path to library source directory (default: src/syntax/library)"
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("docs/library"),
        help="Output directory for generated HTML (default: docs/library)"
    )
    parser.add_argument(
        "--clean",
        action="store_true",
        help="Clean output directory before generating"
    )
    
    args = parser.parse_args()
    
    # Validate source directory
    if not args.src.exists():
        print(f"Error: Source directory not found: {args.src}")
        sys.exit(1)
    
    # Clean output directory if requested
    if args.clean and args.output.exists():
        import shutil
        print(f"Cleaning output directory: {args.output}")
        shutil.rmtree(args.output)
    
    # Create output directory
    args.output.mkdir(parents=True, exist_ok=True)
    
    print("=" * 60)
    print("VB6Parse Library Documentation Generator")
    print("=" * 60)
    print(f"Source: {args.src}")
    print(f"Output: {args.output}")
    print()
    
    # Parse library structure
    print("Parsing library structure...")
    functions, statements = parse_library_structure(args.src)
    
    total_functions = sum(len(cat.items) for cat in functions)
    total_statements = sum(len(cat.items) for cat in statements)
    
    print(f"Found {len(functions)} function categories with {total_functions} items")
    print(f"Found {len(statements)} statement categories with {total_statements} items")
    print()
    
    # Generate main library index
    print("Generating library index...")
    generate_library_index(functions, statements, args.output)
    print()
    
    # Generate function pages
    print("Generating function documentation...")
    for category in functions:
        category_dir = args.output / "functions" / category.slug
        category_dir.mkdir(parents=True, exist_ok=True)
        
        print(f"  Category: {category.display_name} ({len(category.items)} items)")
        generate_category_index(category, "function", category_dir)
        
        for item in category.items:
            generate_item_page(item, category, category_dir)
    print()
    
    # Generate statement pages
    print("Generating statement documentation...")
    for category in statements:
        category_dir = args.output / "statements" / category.slug
        category_dir.mkdir(parents=True, exist_ok=True)
        
        print(f"  Category: {category.display_name} ({len(category.items)} items)")
        generate_category_index(category, "statement", category_dir)
        
        for item in category.items:
            generate_item_page(item, category, category_dir)
    print()
    
    # Generate search index
    print("Generating search index...")
    generate_search_index(functions, statements, args.output)
    print()
    
    print("=" * 60)
    print("✓ Documentation generation complete!")
    print(f"  Total pages: {total_functions + total_statements + len(functions) + len(statements) + 1}")
    print(f"  Output: {args.output}")
    print("=" * 60)


if __name__ == "__main__":
    main()
