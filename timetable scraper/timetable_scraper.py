import aiohttp
import asyncio
from bs4 import BeautifulSoup
from bs4.element import Tag
import pandas as pd
import argparse
import re
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class TimetableScraper:
    def __init__(self):
        self.session = None
        self.df = pd.DataFrame()
        self.course_code = ''
        self.term = ''
        self.year = 0
        self.campus = ''
        self.classes = {}
        self.campuses = {
            "Kensington": 'KENS',
            "Paddington": 'COFA',
            "Canberra City": 'CANC',
            "Canberra ADFA": 'ADFA',
        }

    async def fetch_html(self, url: str):
        if self.session:
            async with self.session.get(url) as response:
                response.raise_for_status()
                content = await response.text()
                return BeautifulSoup(content, 'html.parser')

    async def fetch_timetable(self):
        """
        Fetch timetable data for the given course code and term
        """
        logger.info("Fetching timetable data")
        try:
            html_soup = await self.fetch_html(f'https://timetable.unsw.edu.au/{self.year}/{self.course_code}.html')
        except aiohttp.ClientError as e:
            logger.error(f"Network error: {e}")
            return self.df
        except Exception as e:
            logger.error(f"Error fetching timetable: {e}")
            return self.df
        
        if not html_soup:
            return False

        form_body_cells = html_soup.find_all('td', class_='formBody')
        if not form_body_cells:
            logger.warning("No timetable data found.")
            return self.df

        classes = []
        for cell in form_body_cells:
            inner_cell = cell.find_all('td', class_='formBody', colspan='6')
            if len(inner_cell) == 1 and cell.find('td', class_='data', string='Laboratory'):
                try:
                    class_details, class_term = self._extract_class_details(cell, inner_cell[0]).values()
                except Exception as e:
                    logger.error(f"Error parsing class details: {e}")
                    continue

                if class_term == self.term:
                    classes.append(class_details)

        self.df = pd.DataFrame(classes)

        # logger.info(self.classes) # useful for setups etc

    def _extract_class_details(self, cell, inner_cell):
        """
        Extract class information from a table cell
        """
        rows = cell.find('table').find_all('tr')
        class_info = rows[1].find_all('td', class_='data')
        class_data = rows[3].find_all('td', class_='data')
        class_details = inner_cell.find('table').find_all('tr')[2].find_all('td', class_='data')[0:3]

        self.classes[class_info[0].text.strip()] = class_info[1].text.strip()

        return {
            'class_details': {
                'Class': class_info[0].text.strip(),
                'Section': class_info[1].text.strip(),
                'Status': class_data[1].text.strip(),
                'Enrols/Capacity': class_data[2].text.strip(),
                'Day/Time': f'{class_details[0].text.strip()} {class_details[1].text.strip()}',
                'Location': class_details[2].text.strip(),
            },
            'class_term': class_info[2].text.split()[0].strip()
        }

    def save_timetable_to_csv(self, filename='unsw_timetable.csv'):
        """
        Save timetable data to a CSV file
        """
        self.df.to_csv(filename, index=False)
        logger.info(f"Timetable data saved to {filename}")

    async def check_subject_existence(self, subject):
        logger.info("Checking if subject exists")
        try:
            html_soup = await self.fetch_html('https://timetable.unsw.edu.au/2024/subjectSearch.html')
        except aiohttp.ClientError as e:
            logger.error(f"Network error: {e}")
            return False
        except Exception as e:
            logger.error(f"Error fetching subject list: {e}")
            return False

        if not html_soup:
            return False

        ahrefs = html_soup.find('a', attrs={'name': self.campuses[self.campus]})
        if ahrefs:
            subject_data = ahrefs.find_next('tr')
            if isinstance(subject_data, Tag):
                subjects = [sub.find('td', class_='data').text for sub in subject_data.find_all('tr', class_=['rowHighlight', 'rowLowlight'])]
                return subject in subjects
        return False

    async def check_course_existence(self):
        logger.info("Checking if course exists")
        try:
            html_soup = await self.fetch_html(f'https://timetable.unsw.edu.au/2024/{self.course_code[0:4]}{self.campuses[self.campus]}.html')
        except aiohttp.ClientError as e:
            logger.error(f"Network error: {e}")
            return False
        except Exception as e:
            logger.error(f"Error fetching course list: {e}")
            return False
        
        if not html_soup:
            return False

        classes = []
        categories = html_soup.find_all('td', class_='classSearchSectionHeading')
        for category in categories:
            class_data = category.find_next('tr').find('table')
            if isinstance(class_data, Tag):
                classes.extend([_class.find('td', class_='data').find('a').text for _class in class_data.find_all('tr', class_=['rowHighlight', 'rowLowlight'])])
        return self.course_code in classes

    def course_code_check(self, course_code):
        if not re.match(r'[A-Z]{4}\d{4}', course_code):
            raise argparse.ArgumentTypeError("Course code should have first four letters capital followed by 4 digits")
        return course_code

    async def main(self):
        """
        Set up a database by injecting schema, listing tables, injecting dummy data, and providing details about a specific table
        """
        parser = argparse.ArgumentParser(description="Extracts laboratory schedule from the UNSW Timetable website along with the locations")
        parser.add_argument("--year", type=int, help="Year for which the UNSW Timetable schedule is needed, default value is 2024", default=2024)
        parser.add_argument("--campus", choices=["Kensington", "Paddington", "Canberra City", "Canberra ADFA"], help="Campus Offering for which the UNSW Timetable schedule is needed, default value is Kensington", default='Kensington')
        parser.add_argument("course_code", type=self.course_code_check, help="Course code for which the UNSW Timetable schedule is to be extracted")
        parser.add_argument("term", choices=["T1", "T2", "T3"], help="Term for which the UNSW Timetable schedule is needed")

        namespace = parser.parse_args()
        self.course_code = namespace.course_code
        self.term = namespace.term
        self.year = namespace.year
        self.campus = namespace.campus

        async with aiohttp.ClientSession() as session:
            self.session = session
            if not await self.check_subject_existence(self.course_code[0:4]):
                raise argparse.ArgumentTypeError(f"Subject {self.course_code[0:4]} is not offered at {self.campus} campus")

            if not await self.check_course_existence():
                raise argparse.ArgumentTypeError(f"Course {self.course_code} does not exist at {self.campus} campus")

            await self.fetch_timetable()
            if not self.df.empty:
                self.save_timetable_to_csv()

if __name__ == "__main__":
    scraper = TimetableScraper()
    asyncio.run(scraper.main())
