from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename

import pandas as pd


class ConsistencyCheck:
    """

    """
    def __init__(self):
        pd.options.mode.chained_assignment = None
        self.raw_consistency_report = askopenfilename(title="Select the raw Services Consistency Report")
        self.raw_data = pd.read_excel(
            self.raw_consistency_report,
            sheet_name=None,
            names=[
                "Service Provider",
                "Staff Providing The Service",
                "Service Date",
                "CTID",
                "Service Type",
                "Provider Specific Code"
            ]
        )
        self.staff_list = pd.read_excel(askopenfilename(title="Select the StaffList - All Report"), sheet_name="All")
        self.services = pd.concat(self.raw_data[key] for key in self.raw_data.keys())
        self.re_indexed = self.services.reset_index(drop=True)

        self.services_file = askopenfilename(title="Select the Services List")
        self.raw_services = pd.read_excel(self.services_file).drop("Service You Performed", axis=1)

        self.services = {}
        self.create_services_dict()
        self.drop_list = ["Transitional Housing/Shelter", "Emergency Shelter", "Extreme Cold Weather Shelters"]

    def highlight_null(self, data):
        """
        Highlight all null elements in data-frame red.

        :param data: column data from the data-frame as sent by the apply method
        :return: color parameters to the style method
        """
        blank = pd.isnull(data)
        return ["background-color: red" if value else "" for value in blank]

    def create_services_dict(self):
        """

        :return:
        """
        for row in self.raw_services.index:
            key = self.raw_services.loc[row, "Service Type"]
            self.services[key] = []

        for row in self.raw_services.index:
            value = self.raw_services.loc[row, "Service Provider Specific Code"]
            key = self.raw_services.loc[row, "Service Type"]
            self.services[key].append(value)

        return 1

    def process(self):
        """

        :return:
        """

        data = self.re_indexed[~(self.re_indexed["Service Type"].isin(self.drop_list))].dropna(
            subset=["Staff Providing The Service"]
        ).merge(
            self.staff_list, how="left", left_on="Staff Providing The Service", right_on="CM"
        )
        data["Service Type Errors"] = 0
        data["Provider Specific Service Errors"] = 0
        data["Provider Error"] = 0
        data["Service Date"] = data["Service Date"].dt.date

        # data.style.apply(self.highlight_null)
        for row in data.index:
            service_type = data.loc[row, "Service Type"]
            provider_specific = data.loc[row, "Provider Specific Code"]
            provider = data.loc[row, "Service Provider"]

            try:
                if not(service_type in self.services.keys()):
                    data.loc[row, "Service Type Errors"] += 1
                else:
                    pass
            finally:
                try:
                    if not(provider_specific in self.services[service_type]):
                        data.loc[row, "Provider Specific Service Errors"] += 1
                    else:
                        pass
                except:
                    data.loc[row, "Provider Specific Service Errors"] += 0
                finally:
                    try:
                        if provider == "Transition Projects (TPI) - Agency - SP(19)":
                            data.loc[row, "Provider Error"] += 1
                        else:
                            pass
                    except:
                        data.loc[row, "Provider Error"] += 0
        data["Total Errors"] = data["Provider Error"] + data["Provider Specific Service Errors"] + data["Service Type Errors"]
        self.save_values(data)

    def save_values(self, data_frame):
        """

        :param data_frame:
        :return:
        """
        writer = pd.ExcelWriter(asksaveasfilename(), engine="xlsxwriter")
        for dept_name in list(set(data_frame["Dept"].tolist())):
            dept_df = data_frame[(data_frame["Dept"] == dept_name) & (data_frame["Total Errors"] > 0)]
            dept_df.to_excel(writer, str(dept_name)[:5], engine="xlsxwriter", index=False)
        data_frame.to_excel(writer, "Processed Data", index=False)
        self.re_indexed.to_excel(writer, "Raw Data", index=False)
        writer.save()


if __name__ == "__main__":
    a = ConsistencyCheck()
    a.process()