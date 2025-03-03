from kivy.app import App
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.gridlayout import GridLayout
from kivy.uix.scrollview import ScrollView
from kivy.uix.popup import Popup
from kivy.uix.spinner import Spinner
from kivy.properties import StringProperty
from kivy.core.window import Window
from kivy.uix.screenmanager import Screen
from openpyxl import Workbook
from kivy.uix.widget import Widget
from kivy.graphics import Color, Rectangle
import os

# Custom Date Input Widget with modern color scheme
class DateInput(BoxLayout):
    date = StringProperty("")  # Stores the selected date as a string

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.orientation = "horizontal"
        self.spacing = 5
        self.padding = [10, 10, 10, 10]  # Add padding to the date input

        # Day Spinner
        self.day_spinner = Spinner(
            text="Day",
            values=[str(i) for i in range(1, 32)],
            size_hint=(0.3, None),
            height=40,
            background_normal='',
            background_color=(0.2, 0.6, 0.2, 1),
            color=(1, 1, 1, 1)
        )
        self.day_spinner.bind(text=self.update_date)

        # Month Spinner
        self.month_spinner = Spinner(
            text="Month",
            values=[str(i) for i in range(1, 13)],
            size_hint=(0.3, None),
            height=40,
            background_normal='',
            background_color=(0.2, 0.6, 0.2, 1),
            color=(1, 1, 1, 1)
        )
        self.month_spinner.bind(text=self.update_date)

        # Year Spinner
        self.year_spinner = Spinner(
            text="Year",
            values=[str(i) for i in range(2000, 2060)],
            size_hint=(0.4, None),
            height=40,
            background_normal='',
            background_color=(0.2, 0.6, 0.2, 1),
            color=(1, 1, 1, 1)
        )
        self.year_spinner.bind(text=self.update_date)

        self.add_widget(self.day_spinner)
        self.add_widget(self.month_spinner)
        self.add_widget(self.year_spinner)

    def update_date(self, instance, value):
        # Update the date string when a spinner value changes
        if all(
            [
                self.day_spinner.text != "Day",
                self.month_spinner.text != "Month",
                self.year_spinner.text != "Year",
            ]
        ):
            self.date = f"{self.day_spinner.text}/{self.month_spinner.text}/{self.year_spinner.text}"
        else:
            self.date = ""

# First Screen: Company Details Input
class FirstScreen(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.layout = BoxLayout(orientation='vertical', padding=30, spacing=20)
        self.layout.canvas.before.clear()  # Clear any previous background settings
        with self.layout.canvas.before:
            from kivy.graphics import Color, Rectangle
            Color(0.95, 0.95, 0.95, 1)  # Light gray background color
            self.rect = Rectangle(size=self.layout.size, pos=self.layout.pos)
            self.layout.bind(size=self._update_rect, pos=self._update_rect)

        # Title with black color
        self.layout.add_widget(Label(text="First Form Details", font_size=28, size_hint=(1, 0.1), bold=True, color=(0.2, 0.4, 0.8, 1)))

        # Input fields in a GridLayout
        input_grid = GridLayout(cols=1, spacing=15, size_hint=(1, 0.6), padding=10)
        input_grid.add_widget(Label(text="Company Name", font_size=20, color=(0, 0, 0, 1)))  # Changed color to black
        self.company_name = TextInput(hint_text="Enter Company Name", multiline=False, font_size=18, size_hint_y=None, height=45)
        input_grid.add_widget(self.company_name)

        input_grid.add_widget(Label(text="Address", font_size=20, color=(0, 0, 0, 1)))  # Changed color to black
        self.address = TextInput(hint_text="Enter Address", multiline=False, font_size=18, size_hint_y=None, height=45)
        input_grid.add_widget(self.address)

        input_grid.add_widget(Label(text="Gate Pass Date", font_size=20, color=(0, 0, 0, 1)))  # Changed color to black
        self.gate_pass_date = DateInput()  # Custom DateInput widget
        input_grid.add_widget(self.gate_pass_date)

        input_grid.add_widget(Label(text="Last Job ID", font_size=20, color=(0, 0, 0, 1)))  # Changed color to black
        self.last_job_id = TextInput(hint_text="Enter Last Job ID", multiline=False, font_size=18, size_hint_y=None, height=45)
        input_grid.add_widget(self.last_job_id)

        self.layout.add_widget(input_grid)

        # Next button with a blue color
        next_button = Button(
            text="Next", size_hint=(1, 0.1),
            background_normal='', background_color=(0.2, 0.4, 0.8, 1), font_size=20, height=55,
            color=(1, 1, 1, 1), bold=True
        )
        next_button.bind(on_press=self.next_screen)
        self.layout.add_widget(next_button)

        self.add_widget(self.layout)

    def _update_rect(self, *args):
        self.rect.pos = self.layout.pos
        self.rect.size = self.layout.size

    def next_screen(self, instance):
        # Validate inputs
        if not all(
            [
                self.company_name.text,
                self.address.text,
                self.gate_pass_date.date,
                self.last_job_id.text,
            ]
        ):
            self.show_popup("Error", "Please fill all fields.")
            return

        # Save data and move to the next screen
        app = App.get_running_app()
        app.company_data = {
            "company_name": self.company_name.text,
            "address": self.address.text,
            "gate_pass_date": self.gate_pass_date.date,
            "last_job_id": self.last_job_id.text,
        }
        self.manager.current = "second"

    def show_popup(self, title, message):
        popup = Popup(title=title, size_hint=(0.8, 0.4))
        popup.content = Label(text=message, font_size=20)
        popup.open()


# Second Screen: Instrument Details Input
class SecondScreen(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.layout = BoxLayout(orientation='vertical', padding=30, spacing=20)

        # Set background color of the layout to white
        with self.layout.canvas.before:
            Color(1, 1, 1, 1)  # White background color
            self.rect = Rectangle(size=self.layout.size, pos=self.layout.pos)
            self.layout.bind(size=self._update_rect, pos=self._update_rect)

        # Title with black text color
        self.layout.add_widget(Label(text="Second Form Details", font_size=28, size_hint=(1, 0.1), bold=True, color=(0, 0, 0, 1)))  # Black text color

        # Display company data with black text color
        self.company_info = Label(font_size=16, size_hint=(1, 0.1), color=(0, 0, 0, 1))  # Black text color
        self.layout.add_widget(self.company_info)

        # Create a ScrollView for the input fields
        scroll_view = ScrollView(size_hint=(1, None), height=400)  # Adjust height for scroll area

        # Create GridLayout to hold the form fields
        input_grid = GridLayout(cols=1, spacing=15, size_hint_y=None, padding=10)
        input_grid.bind(minimum_height=input_grid.setter('height'))

        self.fields = [
            "sr_no", "name_of_instrument", "make", "model", "id_no", "sr_num",
            "low_range", "high_range", "unit", "l_c", "accuracy", "location",
            "calibration_date", "certificate_no", "due_date", "remarks"  # Interchange due_date and certificate_no
        ]
        self.inputs = {}
        for i, field in enumerate(self.fields):
            # Label for each field with black text color
            input_grid.add_widget(Label(text=field.replace("_", " ").title(), font_size=18, color=(0, 0, 0, 1)))  # Black text color

            if field == "calibration_date":
                # Use DateInput for calibration_date
                self.inputs[field] = DateInput(height=55, size_hint_y=None, padding=(5, 5, 5, 5))  # Set height and padding for DateInput
                input_grid.add_widget(self.inputs[field])
                # Add space after the date fields
                input_grid.add_widget(Widget(size_hint_y=None, height=10))  # Space after date fields
                
            elif field == "certificate_no":
                # Use TextInput for certificate_no field
                self.inputs[field] = TextInput(hint_text=f"Enter {field.replace('_', ' ')}", multiline=False, font_size=16, size_hint_y=None, height=45)
                input_grid.add_widget(self.inputs[field])
                # Add space after certificate_no
                input_grid.add_widget(Widget(size_hint_y=None, height=10))  # Space after certificate_no

            elif field == "due_date":
                # Use DateInput for due_date
                self.inputs[field] = DateInput(height=55, size_hint_y=None, padding=(5, 5, 5, 5))  # Set height and padding for DateInput
                input_grid.add_widget(self.inputs[field])
                # Add space after the date fields
                input_grid.add_widget(Widget(size_hint_y=None, height=10))  # Space after date fields

            elif field == "remarks":
                # Use TextInput for remarks field
                self.inputs[field] = TextInput(hint_text=f"Enter {field.replace('_', ' ')}", multiline=False, font_size=16, size_hint_y=None, height=45)
                input_grid.add_widget(self.inputs[field])
                # Add space after remarks
                input_grid.add_widget(Widget(size_hint_y=None, height=10))  # Space after remarks
            else:
                # Use TextInput for other fields
                self.inputs[field] = TextInput(hint_text=f"Enter {field.replace('_', ' ')}", multiline=False, font_size=16, size_hint_y=None, height=45)
                input_grid.add_widget(self.inputs[field])

        # Add GridLayout with input fields to the ScrollView
        scroll_view.add_widget(input_grid)

        # Add ScrollView to the main layout
        self.layout.add_widget(scroll_view)

        # Add and Save buttons with spacing between
        button_layout = BoxLayout(size_hint=(1, 0.1), spacing=10)
        add_button = Button(text="Add Data", background_color=(0, 0.7, 0, 1), font_size=18)
        add_button.bind(on_press=self.add_instrument)
        save_button = Button(text="Download", background_color=(0.2, 0.6, 1, 1), font_size=18)
        save_button.bind(on_press=self.save_data)
        button_layout.add_widget(add_button)
        button_layout.add_widget(save_button)
        self.layout.add_widget(button_layout)

        self.add_widget(self.layout)
        self.instruments = []

    def _update_rect(self, *args):
        self.rect.pos = self.layout.pos
        self.rect.size = self.layout.size

    def on_enter(self):
        # Display company data with black text color
        app = App.get_running_app()
        self.company_info.text = (
            f"Company: {app.company_data['company_name']}\n"
            f"Address: {app.company_data['address']}\n"
            f"Gate Pass Date: {app.company_data['gate_pass_date']}\n"
            f"Last Job ID: {app.company_data['last_job_id']}"
        )

    def add_instrument(self, instance):
        # Validate inputs
        if not all(
            self.inputs[field].text if isinstance(self.inputs[field], TextInput) else self.inputs[field].date
            for field in self.fields
        ):
            self.show_popup("Error", "Please fill all fields.")
            return

        # Save instrument data
        instrument = {}
        for field in self.fields:
            if isinstance(self.inputs[field], DateInput):
                instrument[field] = self.inputs[field].date
            else:
                instrument[field] = self.inputs[field].text
        self.instruments.append(instrument)

        # Clear inputs
        for field in self.fields:
            if isinstance(self.inputs[field], DateInput):
                self.inputs[field].date = ""
            else:
                self.inputs[field].text = ""

        self.show_popup("Success", "Instrument added successfully.")

    def save_data(self, instance):
        if not self.instruments:
            self.show_popup("Error", "No instruments added.")
            return

        # Export to Excel
        wb = Workbook()
        ws = wb.active
        ws.title = "Instrument Data"

        # Write company data with black text color
        app = App.get_running_app()
        ws.append(["Company Name", app.company_data["company_name"]])
        ws.append(["Address", app.company_data["address"]])
        ws.append(["Gate Pass Date", app.company_data["gate_pass_date"]])
        ws.append(["Last Job ID", app.company_data["last_job_id"]])
        ws.append([])

        # Write instrument headers
        ws.append(self.fields)

        # Write instrument data
        for instrument in self.instruments:
            ws.append([instrument[field] for field in self.fields])

        # Save file
        file_path = "instrument_data.xlsx"
        wb.save(file_path)
        self.show_popup("Success", f"Data exported to {file_path}")

    def show_popup(self, title, message):
        popup = Popup(title=title, size_hint=(0.8, 0.4))
        popup.content = Label(text=message, font_size=18, color=(0, 0, 0, 1))  # Black text color for popup
        popup.open()


# App Manager
class MyApp(App):
    def build(self):
        Window.size = (360, 640)
        sm = ScreenManager()
        sm.add_widget(FirstScreen(name="first"))
        sm.add_widget(SecondScreen(name="second"))
        return sm


if __name__ == "__main__":
    MyApp().run()
