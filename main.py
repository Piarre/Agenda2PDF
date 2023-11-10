import icalendar
import jinja2
import pdfkit
import os
from datetime import datetime

__WKHTMLTOPDF_PATH__ = "C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe"
__FILE_NAME__ = "events.ics"

output_folder = os.path.join("out", __FILE_NAME__)
os.makedirs(output_folder, exist_ok=True)

with open(__FILE_NAME__, mode='r', encoding='utf-8') as f:
    calendar = icalendar.Calendar.from_ical(f.read())
    for event in calendar.walk('VEVENT'):

        __SUMMARY__ = str(event.get("SUMMARY"))
        __DESCRIPTION__ = event.get("DESCRIPTION")
        __ADDRESSE__ = event.get("LOCATION")
        __UNFORMATED_DATE__ = event.get("CREATED").dt 

        formatted_date = __UNFORMATED_DATE__.strftime("%d/%m/%Y")

        context = {'SUMMARY': __SUMMARY__}

        if __DESCRIPTION__ is not None:
            context['description'] = __DESCRIPTION__

        if formatted_date is not None:
            context['fulldate'] = formatted_date

        if __ADDRESSE__ is not None:
            context['addresse'] = __ADDRESSE__

        template_loader = jinja2.FileSystemLoader('./template')
        template_env = jinja2.Environment(loader=template_loader)

        html_template = 'template.html'
        template = template_env.get_template(html_template)
        output_text = template.render(context)

        config = pdfkit.configuration(wkhtmltopdf=__WKHTMLTOPDF_PATH__)
        __OUTPUT_FILE_NAME__ = f'{str(event.get("CREATED").dt).replace(" ", "_").replace(".", "_").replace("\n", "_").replace("/", "_").replace(":", "_").replace("(", "_").replace(")", "_").replace("\\", "_").replace("|", "_").replace("?", "_").replace("*", "_").replace("<", "_").replace(">", "_").replace("=", "_") + str(event.get("SUMMARY")).replace(" ", "_").replace(".", "_").replace("/", "_").replace(":", "_").replace("(", "_").replace(")", "_").replace("\\", "_").replace("|", "_").replace("?", "_").replace("*", "_").replace("<", "_").replace(">", "_").replace("=", "_").replace("\n", "_")}.pdf'
        print(__OUTPUT_FILE_NAME__)
        
        output_pdf = os.path.join(output_folder, __OUTPUT_FILE_NAME__)

        pdfkit.from_string(output_text, output_pdf, configuration=config, css="style.css", options={'encoding': 'utf-8'})
