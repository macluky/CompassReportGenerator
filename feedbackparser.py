import openpyxl


class FeedbackParser:

    def __init__(self, _file="/Users/macluky/Downloads/feedback_responses.xlsx"):
        self.file = _file
        self.professionals = dict()

        workbook = openpyxl.load_workbook(self.file, read_only=False)
        sheet = workbook.active

        for row in range(2, (sheet.max_row + 1)):
            responses = dict()
            name = None
            # extract the feedback from the feedback giver
            for column in range(1, sheet.max_column):
                key = sheet.cell(row=1, column=column).value
                value = sheet.cell(row=row, column=column).value
                responses[key] = value
                if key == "email":
                    name = value

            # find or create the professional
            if name in self.professionals.keys():
                professional = self.professionals[name]
            else:
                professional = []
                self.professionals[name] = professional

            # add the responses
            professional.append(responses)

    def responses_for_email(self, email):
        if email in self.professionals.keys():
            return self.professionals[email]
        else:
            print("found no feedback for: " + email)
            return None


if __name__ == '__main__':
    fp = FeedbackParser()
    fp.responses_for_email("s@b.de")
    rp = fp.responses_for_email("a.ganga@topdesk.com")
