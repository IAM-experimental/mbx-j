from bs4 import BeautifulSoup, NavigableString
import re


def html_to_jira_markup(html_content: str) -> str:
    """
    Convert HTML content to JIRA markup format.
    
    Args:
        html_content: HTML string to convert
        
    Returns:
        JIRA-formatted markup string
    """
    soup = BeautifulSoup(html_content, 'html.parser')
    
    def process_element(element, list_type=None, list_level=0):
        """Recursively process HTML elements and convert to JIRA markup."""
        if isinstance(element, NavigableString):
            text = str(element)
            # Preserve whitespace but clean up excessive newlines
            if text.strip():
                return text
            return ''
        
        result = []
        tag = element.name
        
        # Headers
        if tag in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
            level = tag[1]
            content = ''.join(process_element(child) for child in element.children)
            result.append(f'h{level}. {content.strip()}\n\n')
        
        # Bold
        elif tag in ['b', 'strong']:
            content = ''.join(process_element(child) for child in element.children)
            result.append(f'*{content}*')
        
        # Italic
        elif tag in ['i', 'em']:
            content = ''.join(process_element(child) for child in element.children)
            result.append(f'_{content}_')
        
        # Underline
        elif tag == 'u':
            content = ''.join(process_element(child) for child in element.children)
            result.append(f'+{content}+')
        
        # Strikethrough
        elif tag in ['s', 'strike', 'del']:
            content = ''.join(process_element(child) for child in element.children)
            result.append(f'-{content}-')
        
        # Code
        elif tag == 'code':
            content = ''.join(process_element(child) for child in element.children)
            result.append(f'{{{content}}}')
        
        # Preformatted/Code blocks
        elif tag == 'pre':
            content = element.get_text()
            result.append(f'{{code}}\n{content}\n{{code}}\n\n')
        
        # Links
        elif tag == 'a':
            href = element.get('href', '')
            content = ''.join(process_element(child) for child in element.children)
            if href:
                result.append(f'[{content.strip()}|{href}]')
            else:
                result.append(content)
        
        # Images
        elif tag == 'img':
            src = element.get('src', '')
            alt = element.get('alt', '')
            if src:
                result.append(f'!{src}!')
        
        # Paragraphs
        elif tag == 'p':
            content = ''.join(process_element(child) for child in element.children)
            result.append(f'{content.strip()}\n\n')
        
        # Line breaks
        elif tag == 'br':
            result.append('\n')
        
        # Horizontal rule
        elif tag == 'hr':
            result.append('----\n\n')
        
        # Blockquote
        elif tag == 'blockquote':
            content = ''.join(process_element(child) for child in element.children)
            lines = content.strip().split('\n')
            for line in lines:
                if line.strip():
                    result.append(f'bq. {line.strip()}\n')
            result.append('\n')
        
        # Lists
        elif tag == 'ul':
            for child in element.children:
                if child.name == 'li':
                    result.append(process_list_item(child, '*', list_level))
        
        elif tag == 'ol':
            for child in element.children:
                if child.name == 'li':
                    result.append(process_list_item(child, '#', list_level))
        
        # Tables
        elif tag == 'table':
            result.append(process_table(element))
        
        # Divs and spans - process children
        elif tag in ['div', 'span', 'body', 'html']:
            for child in element.children:
                result.append(process_element(child, list_type, list_level))
        
        # Other elements - just process children
        else:
            for child in element.children:
                result.append(process_element(child, list_type, list_level))
        
        return ''.join(result)
    
    def process_list_item(li, marker, level):
        """Process list item with proper nesting."""
        prefix = marker * (level + 1)
        content_parts = []
        
        for child in li.children:
            if child.name in ['ul', 'ol']:
                # Nested list
                new_marker = '*' if child.name == 'ul' else '#'
                for nested_li in child.find_all('li', recursive=False):
                    content_parts.append('\n' + process_list_item(nested_li, new_marker, level + 1))
            else:
                content_parts.append(process_element(child, marker, level))
        
        content = ''.join(content_parts).strip()
        return f'{prefix} {content}\n'
    
    def process_table(table):
        """Convert HTML table to JIRA table markup."""
        rows = []
        
        # Process header rows
        for thead in table.find_all('thead'):
            for tr in thead.find_all('tr'):
                cells = []
                for th in tr.find_all(['th', 'td']):
                    content = ''.join(process_element(child) for child in th.children).strip()
                    cells.append(content)
                if cells:
                    rows.append('||' + '||'.join(cells) + '||')
        
        # Process body rows
        for tbody in table.find_all('tbody'):
            for tr in tbody.find_all('tr'):
                cells = []
                for td in tr.find_all(['td', 'th']):
                    content = ''.join(process_element(child) for child in td.children).strip()
                    cells.append(content)
                if cells:
                    rows.append('|' + '|'.join(cells) + '|')
        
        # If no thead/tbody, just process all rows
        if not rows:
            for tr in table.find_all('tr', recursive=False):
                cells = []
                has_th = bool(tr.find('th'))
                for cell in tr.find_all(['td', 'th']):
                    content = ''.join(process_element(child) for child in cell.children).strip()
                    cells.append(content)
                if cells:
                    if has_th:
                        rows.append('||' + '||'.join(cells) + '||')
                    else:
                        rows.append('|' + '|'.join(cells) + '|')
        
        return '\n'.join(rows) + '\n\n' if rows else ''
    
    # Process the HTML
    jira_markup = process_element(soup)
    
    # Clean up excessive newlines (more than 2 consecutive)
    jira_markup = re.sub(r'\n{3,}', '\n\n', jira_markup)
    
    return jira_markup.strip()


def _example_usage():
    """Example usage of the converter."""
    sample_html = """
    <h1>Bug Report</h1>
    <p>This is a <strong>critical</strong> issue with <em>high priority</em>.</p>
    <h2>Steps to Reproduce</h2>
    <ol>
        <li>Open the application</li>
        <li>Click on <code>Settings</code></li>
        <li>Navigate to <a href="https://example.com">Profile</a></li>
    </ol>
    <h2>Expected vs Actual</h2>
    <ul>
        <li>Expected: User profile loads</li>
        <li>Actual: <s>Error message</s> Application crashes</li>
    </ul>
    <blockquote>This was reported by multiple users</blockquote>
    <pre>Error: NullPointerException
    at line 42</pre>
    """
    
    print(html_to_jira_markup(sample_html))


if __name__ == '__main__':
    _example_usage()
