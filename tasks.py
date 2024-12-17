from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF

import os
import logging
from dotenv import load_dotenv
from transitions import Machine

# --------------------------
# CONFIGURATIONS
# --------------------------
# Load credentials securely
load_dotenv('login_credentials.env')
USERNAME = os.getenv("RSB_ROBOT_USERNAME")
PASSWORD = os.getenv("RSB_ROBOT_PASSWORD")

# Logging configuration
logging.basicConfig(
    filename="robot_spare_bin.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)

# --------------------------
# STATE TRANSITION MACHINE
# --------------------------
class BotStateMachine:
    """State machine for the bot workflow."""

    states = ["start", "login", "process_data", "export", "log_out", "error", "completed"]

    def __init__(self):
        self.machine = Machine(model=self, states=BotStateMachine.states, initial="start")

        # Transitions
        self.machine.add_transition("proceed_to_login", "start", "login")
        self.machine.add_transition("process_sales_data", "login", "process_data")
        self.machine.add_transition("export_results", "process_data", "export")
        self.machine.add_transition("proceed_to_log_out", "export", "log_out")
        self.machine.add_transition("complete_task", "log_out", "completed")
        self.machine.add_transition("handle_error", "*", "error")
        self.machine.add_transition("recover", "error", "start")


class RobotSpareBinBot:
    """Main bot class for handling tasks with state transitions and exceptions."""

    def __init__(self):
        self.state_machine = BotStateMachine()

    def run(self):
        """Run the bot with state transitions and exception handling."""
        try:
            # Start state
            logging.info("Starting the bot...")
            self.state_machine.proceed_to_login()

            # Login state
            self.open_the_intranet_website()
            self.log_in()
            self.state_machine.process_sales_data()

            # Process data state
            self.download_excel_file()
            self.fill_form_with_excel_data()
            self.state_machine.export_results()

            # Export state
            self.collect_results()
            self.export_as_pdf()
            self.state_machine.proceed_to_log_out()

            # Log out state
            self.log_out()
            self.state_machine.complete_task()
            logging.info("Bot completed successfully.")

        except Exception as e:
            logging.error(f"Unexpected error occurred: {e}")
            self.state_machine.handle_error()

    # --------------------------
    # Bot Functionalities
    # --------------------------
    def open_the_intranet_website(self):
        """Navigates to the given URL."""
        try:
            browser.configure(slowmo=100)
            browser.goto("https://robotsparebinindustries.com/")
            logging.info("Navigated to intranet website.")
        except Exception as e:
            logging.error(f"Error opening the website: {e}")
            raise

    def log_in(self):
        """Logs in securely using credentials."""
        try:
            page = browser.page()
            page.fill("#username", USERNAME)
            page.fill("#password", PASSWORD)
            page.click("button:text('Log in')")
            logging.info("Logged in successfully.")
        except Exception as e:
            logging.error(f"Login failed: {e}")
            raise

    def download_excel_file(self):
        """Downloads the Excel file from a given URL."""
        try:
            http = HTTP()
            http.download(url="https://robotsparebinindustries.com/SalesData.xlsx", overwrite=True)
            logging.info("Excel file downloaded successfully.")
        except Exception as e:
            logging.error(f"Failed to download Excel file: {e}")
            raise

    def fill_and_submit_sales_form(self, sales_rep):
        """Fills and submits sales data for one row."""
        try:
            page = browser.page()
            page.fill("#firstname", sales_rep["First Name"])
            page.fill("#lastname", sales_rep["Last Name"])
            page.select_option("#salestarget", str(sales_rep["Sales Target"]))
            page.fill("#salesresult", str(sales_rep["Sales"]))
            page.click("text=Submit")
            logging.info(f"Sales data submitted for: {sales_rep['First Name']} {sales_rep['Last Name']}")
        except Exception as e:
            logging.warning(f"Error submitting sales form: {e}")

    def fill_form_with_excel_data(self):
        """Reads Excel data and fills out the sales form."""
        try:
            excel = Files()
            excel.open_workbook("SalesData.xlsx")
            worksheet = excel.read_worksheet_as_table("data", header=True)
            for row in worksheet:
                self.fill_and_submit_sales_form(row)
            excel.close_workbook()
        except Exception as e:
            logging.error(f"Error processing Excel file: {e}")
            raise

    def collect_results(self):
        """Takes a screenshot of the sales results."""
        try:
            page = browser.page()
            page.screenshot(path="output/sales_summary.png")
            logging.info("Screenshot captured.")
        except Exception as e:
            logging.error(f"Error taking screenshot: {e}")
            raise

    def export_as_pdf(self):
        """Exports sales results to PDF."""
        try:
            page = browser.page()
            sales_results_html = page.locator("#sales-results").inner_html()
            pdf = PDF()
            pdf.html_to_pdf(sales_results_html, "output/sales_results.pdf")
            logging.info("Sales results exported to PDF.")
        except Exception as e:
            logging.error(f"Error exporting PDF: {e}")
            raise

    def log_out(self):
        """Logs out from the application."""
        try:
            page = browser.page()
            page.click("text=Log out")
            logging.info("Logged out successfully.")
        except Exception as e:
            logging.error(f"Error during log out: {e}")
            raise


@task
def robot_spare_bin_python():
    """Task to execute the Robot Spare Bin Bot."""
    bot = RobotSpareBinBot()
    bot.run()
