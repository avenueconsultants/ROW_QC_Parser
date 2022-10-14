class RW53:
    def __init__(self, name, page_numbers) -> None:
        self.name = 'RW-53'
        self.page_numbers = page_numbers
        self.prepared_by = None
        self.prepared_by_date = None
        self.pin = None
        self.project_number = None
        self.region = None
        self.county = None
        self.routes = None
        self.project_name = name
        self.summary_number = None 
        self.parcels = None

class SummaryPage:
    def __init__(self, title, page_numbers) -> None:
        self.title = title
        self.page_numbers = page_numbers
        self.pin = None
        self.project_number = None
        self.project_name = None
        self.region = None
        self.county = None
        self.routes = None
        self.abc = None
        self.consultant = None
        self.prepared_date = None
        self.form_date = None
        self.parcels = None

class Parcel:
    def __init__(self) -> None:
        self.parcel_number = None
        self.grantor_name_short = None
        self.square_feet = None
        self.acres = None
        self.deed_type = None
        self.is_void = False
        self.is_void_and_replace = False
        self.map_sheets = None
        self.notes = None

        # self.deed_searched_by = None
        # self.deed_searched_by_date = None
        # self.pin = None
        # self.project_number = None
        # self.summary_number = None
        # self.ownership_number = None
        # self.county = None
        # self.Tax