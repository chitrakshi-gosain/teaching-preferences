import requests
from bs4 import BeautifulSoup
import pandas as pd

class TimetableScraper:
    def __init__(self, course_code, term):
        self.course_code = course_code
        self.term = term

    def fetch_timetable(self):
        """
        Fetch timetable data for the given course code and term
        """
        url = f'https://timetable.unsw.edu.au/2024/{self.course_code}.html'
        response = requests.get(url)
        response.raise_for_status()

        html_soup = BeautifulSoup(response.content, 'html.parser')

        form_body_cells = html_soup.find_all('td', class_='formBody')
        if not form_body_cells:
            print("No td tags found with class 'formBody'")
            return pd.DataFrame()

        classes = []
        for cell in form_body_cells:
            inner_cell = cell.find_all('td', class_='formBody', colspan='6')
            if len(inner_cell) == 1 and cell.find('td', class_='data', string='Laboratory'):
                class_info = self._extract_class_info(cell, inner_cell[0])
                current_term = class_info.get('Term', '').split()[0]
                if current_term == self.term:
                    class_info['Term'] = current_term
                    classes.append(class_info)

        return pd.DataFrame(classes)

    def _extract_class_info(self, cell, inner_cell):
        """
        Extract class information from a table cell
        """
        class_info = cell.find('table').find_all('tr')[1].find_all('td', class_='data')
        class_details = inner_cell.find('table').find_all('tr')[2].find_all('td', class_='data')[0:3]
        return {
            'Class Number': class_info[0].text,
            'Class Code': class_info[1].text,
            'Term': class_info[2].text,
            'Day': class_details[0].text,
            'Time': class_details[1].text,
            'Location': class_details[2].text,
        }

    def save_timetable_to_csv(self, df, filename='unsw_timetable.csv'):
        """
        Save timetable data to a CSV file
        """
        df.to_csv(filename, index=False)

if __name__ == "__main__":
    course_code = input("Enter the course code: ")
    term = input("Enter the term: ")
    scraper = TimetableScraper(course_code, term)
    timetable_data = scraper.fetch_timetable()
    if not timetable_data.empty:
        scraper.save_timetable_to_csv(timetable_data)
