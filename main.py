import json
from docx import Document
import os
import re

def clean_text(text):
    """Remove markdown-like formatting and extra whitespace."""
    text = re.sub(r'[\*_]+', '', text)  # Remove markdown bold/italic
    text = re.sub(r'\[.*?\]\((.*?)\)', r'\1', text)  # Extract URL from [text](url)
    return text.strip()

def extract_url(text):
    """Extract complete URLs from text, handling split lines and markdown."""
    markdown_match = re.search(r'\[.*?\]\((.*?)\)', text)
    if markdown_match:
        return markdown_match.group(1)
    
    url_match = re.search(r'https?://[^\s>"\'\]\)]+', text)
    return url_match.group(0) if url_match else None

def doc_to_json(doc_path, json_path):
    """
    Improved version to handle:
    - Tool Name detection without ### prefix
    - Accurate field detection with values on the same line
    - Multi-line content with subheadings
    - Multiple URLs in Screenshots
    - Ensure Product Type is populated
    """
    if not os.path.exists(doc_path):
        print(f"Error: File {doc_path} not found")
        return

    try:
        document = Document(doc_path)
        all_tools = []
        current_tool = None
        current_key = None
        current_subsection = None
        in_full_description = False
        in_screenshots = False
        partial_url = None
        expect_tool_name_value = False

        # All required fields
        FIELDS = [
            "Tool Name", "Website Link", "Logo", "Screenshots", "Short Description",
            "Full Description", "Slug", "Meta Title", "Meta Description",
            "Category", "Product Type", "Tags"
        ]

        ARRAY_FIELDS = ["Category", "Tags", "Screenshots"]
        URL_FIELDS = ["Website Link", "Logo", "Screenshots"]

        # Subsections under Full Description
        FULL_DESCRIPTION_SUBSECTIONS = [
            "Introduction", "Key Features", "Use Cases", "How It Works",
            "Why Choose", "Future Vision", "Conclusion"
        ]

        def is_new_tool(line):
            """Determine if the line indicates the start of a new tool."""
            cleaned_line = clean_text(line)
            return bool(re.match(r'^(###\s*)?Tool Name\s*:$', cleaned_line, re.IGNORECASE))

        for para in document.paragraphs:
            line = para.text.strip()
            if not line:
                continue

            cleaned_line = clean_text(line)
            print(f"Processing line: {line}")

            if is_new_tool(line):
                print(f"Detected new tool header: {cleaned_line}")
                if current_tool and current_tool.get("Tool Name"):
                    print(f"Adding tool: {current_tool['Tool Name']}")
                    all_tools.append(current_tool)
                current_tool = {field: [] if field in ARRAY_FIELDS else "" for field in FIELDS}
                current_tool["Full Description"] = {sub: "" for sub in FULL_DESCRIPTION_SUBSECTIONS}
                current_key = "Tool Name"
                current_subsection = None
                in_full_description = False
                in_screenshots = False
                partial_url = None
                expect_tool_name_value = True
                continue

            if expect_tool_name_value:
                current_tool["Tool Name"] = cleaned_line
                print(f"Set Tool Name: {cleaned_line}")
                expect_tool_name_value = False
                continue

            if current_tool is None:
                print(f"Skipping line (no tool initialized): {line}")
                continue

            if partial_url:
                if re.match(r'https?://', line):
                    if current_key == "Screenshots":
                        current_tool["Screenshots"].append(partial_url)
                    else:
                        current_tool[current_key] = partial_url
                    partial_url = extract_url(cleaned_line)
                else:
                    partial_url += " " + cleaned_line
                    complete_url = extract_url(partial_url)
                    if complete_url:
                        if current_key == "Screenshots":
                            current_tool["Screenshots"].append(complete_url)
                        else:
                            current_tool[current_key] = complete_url
                        partial_url = None
                continue

            # Detect field (allow value after colon on the same line)
            field_match = re.match(r'^(###\s*)?(.*?)\s*:\s*(.*)?$', line) or re.match(r'^(##\s*)?(.*?)\s*:\s*(.*)?$', line)
            if field_match:
                field_candidate = clean_text(field_match.group(2).strip())
                matched_field = next((f for f in FIELDS if f.lower() == field_candidate.lower()), None)
                value = field_match.group(3).strip() if field_match.group(3) else ""

                if matched_field:
                    print(f"Detected field: {matched_field} with value: {value}")
                    current_key = matched_field
                    in_full_description = (current_key == "Full Description")
                    in_screenshots = (current_key == "Screenshots")
                    current_subsection = None
                    if value:
                        if current_key in ARRAY_FIELDS:
                            current_tool[current_key] = [v.strip() for v in value.split(",") if v.strip()]
                        elif current_key in URL_FIELDS:
                            url = extract_url(value)
                            if url:
                                if current_key == "Screenshots":
                                    current_tool["Screenshots"] = [url]
                                else:
                                    current_tool[current_key] = url
                            elif value:
                                partial_url = value
                        else:
                            current_tool[current_key] = value
                    continue

            if line.lower().startswith("logo:"):
                print("Detected Logo field")
                current_key = "Logo"
                in_screenshots = False
                in_full_description = False
                current_subsection = None
                value = cleaned_line.split(":", 1)[1].strip()
                url = extract_url(value)
                if url:
                    current_tool["Logo"] = url
                elif value:
                    partial_url = value
                continue

            if line.lower().startswith("screenshots:"):
                print("Detected Screenshots field")
                current_key = "Screenshots"
                in_screenshots = True
                in_full_description = False
                current_subsection = None
                value = cleaned_line.split(":", 1)[1].strip()
                url = extract_url(value)
                if url:
                    current_tool["Screenshots"].append(url)
                elif value:
                    partial_url = value
                continue

            if in_screenshots:
                url = extract_url(cleaned_line)
                if url:
                    current_tool["Screenshots"].append(url)
                    print(f"Added screenshot URL: {url}")
                elif re.match(r'https?://', cleaned_line):
                    partial_url = cleaned_line
                else:
                    in_screenshots = False
                continue

            if in_full_description:
                subsection_match = re.match(r'^(####\s*)?(.*?)\s*:\s*(.*)?$', line)
                if subsection_match:
                    subsection_candidate = clean_text(subsection_match.group(2).strip())
                    matched_subsection = next(
                        (s for s in FULL_DESCRIPTION_SUBSECTIONS if subsection_candidate.lower().startswith(s.lower())),
                        None
                    )
                    value = subsection_match.group(3).strip() if subsection_match.group(3) else ""
                    if matched_subsection:
                        print(f"Detected subsection: {matched_subsection}")
                        current_subsection = matched_subsection
                        current_tool["Full Description"][current_subsection] = value
                        continue
                if current_subsection:
                    current_tool["Full Description"][current_subsection] += " " + cleaned_line
                continue

            if current_key and not in_screenshots and not in_full_description:
                if current_key in ARRAY_FIELDS:
                    current_tool[current_key].extend([v.strip() for v in cleaned_line.split(",") if v.strip()])
                elif current_key in URL_FIELDS:
                    if not current_tool[current_key]:
                        url = extract_url(cleaned_line)
                        if url:
                            if current_key == "Screenshots":
                                current_tool["Screenshots"].append(url)
                            else:
                                current_tool[current_key] = url
                        else:
                            partial_url = cleaned_line
                else:
                    current_tool[current_key] += " " + cleaned_line

        # Add the last tool
        if current_tool and current_tool.get("Tool Name"):
            if partial_url and current_key in URL_FIELDS:
                if current_key == "Screenshots":
                    current_tool["Screenshots"].append(partial_url)
                else:
                    current_tool[current_key] = partial_url
            print(f"Adding final tool: {current_tool['Tool Name']}")
            all_tools.append(current_tool)

        # Post-processing
        for tool in all_tools:
            for field in ARRAY_FIELDS:
                if field in tool:
                    tool[field] = list(set(filter(None, tool[field])))
            
            for field in URL_FIELDS:
                if field in tool:
                    if isinstance(tool[field], list):
                        tool[field] = [extract_url(url) or url for url in tool[field]]
                    elif isinstance(tool[field], str):
                        tool[field] = extract_url(tool[field]) or tool[field]

            for subsection in tool["Full Description"]:
                tool["Full Description"][subsection] = clean_text(tool["Full Description"][subsection]).strip()

            # Ensure Product Type has a default value if empty
            if not tool["Product Type"]:
                tool["Product Type"] = "AI Tool"

            # Ensure Category is populated based on context if empty
            if not tool["Category"]:
                if "Writing" in tool["Short Description"] or "Writing" in tool["Full Description"]["Introduction"]:
                    tool["Category"] = ["AI Writing Assistant"]
                elif "Image" in tool["Short Description"] or "Image" in tool["Full Description"]["Introduction"]:
                    tool["Category"] = ["AI Image Generation Tool"]
                else:
                    tool["Category"] = ["AI Tool"]

        # Save JSON
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(all_tools, f, indent=4, ensure_ascii=False)

        print(f"Successfully processed {len(all_tools)} tools to {json_path}")

    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    doc_to_json("ai.docx", "tools_output.json")