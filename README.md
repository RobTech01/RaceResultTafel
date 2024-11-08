
# Digitale Ergebnistafel für Leichtathletik-Wettkämpfe (Proof-Of-Concept)

## Description
**Digitale Ergebnistafel für Leichtathletik-Wettkämpfe** is a Python-based automation tool designed to create dynamic digital result boards for track events. Leveraging HTML web scraping, it extracts essential information from the official DLV results page for most national events. The tool automizes the process of displaying lane information for live-stream purposes in real-time using PowerPoint. This is a proof-of-concept.

## Motivation
The project was initiated to graphically display lane occupations and rankings for track events, enhancing the presentation of live-streamed track events.

## Installation
- **Requirements**: Python 3.x is required.
- **Dependencies**: Install all necessary dependencies using the provided `requirements.txt` file with the command:  
  `pip install -r requirements.txt`.
- **PowerPoint Template**: Ensure the package is in the same folder as `main.py` or as specified in the code.

## Usage
To use the script, have ONE active PowerPoint file open (in Presenter mode f5) and check the box for developer options in the settings. Deactivate the permissions once finished.

Note: The URL for the results page is currently hardcoded into the Python code. If this tools prooves useful / needed, future updates will allow specifying the URL directly via a CLI command.

## PowerPoint Design Requirements
- Colors and elements within the PowerPoint slides are customizable.
- Elements can be added but should not be grouped.
- The structure of Slide 1 and 2 must remain consistent to ensure proper functionality.
- Editing is allowed, within the scope of the rules outlined above.

## Features and Benefits
- **Ease of Redesign and Adjustment**: Utilizes PowerPoint for straightforward customization.
- **Autonomous Operation**: Once initiated, it runs autonomously to scrape data and populate the presentation template.
- **Iterative Updates**: Designed to iteratively update presentations with new data as it becomes available.

## Roadmap and Future Plans
- Introduce a CLI command for specifying the URL of the results page.
- Improve the duplication process to eliminate visual glitches by cloning data fields outside the viewport.

## License
This project is licensed under the **GNU General Public License v3.0**. A copy of the license can be found in the LICENSE file within the project repository.

## Contact Information
For support, questions, or collaboration opportunities, please contact via GitHub.
