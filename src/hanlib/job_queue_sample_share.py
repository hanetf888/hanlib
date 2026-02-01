import json

def sample_1():
    table: str = "data.rundeck_job_queue"
    job_type: str = "graph_email_send_uat"
    json_payload: dict = {
        "to": "mark.tan@hanetf.com",
        "subject": "Test subject",
        "cc": ["karen.kiwanuka@hanetf.com"],
        "bcc": ["ramon.williams@hanetf.com"],
        "body_text": "Test body",
        "body_html": None,
        "attachments": [r"Y:\dev-geoffrey-mark\geoffrey\tests\testdata\test_copy_files_to_sharepoint\Bar_list-RMAU-XS2115336336-all-all.xlsx"],
        "draft_or_send": "send"
    }

    row_data: dict = {"job_type": job_type, "json_payload": json_payload}
    print("sample_1 row_data: {}".format(json.dumps(row_data)))

def sample_2():
    table: str = "data.rundeck_job_queue"
    job_type: str = "graph_email_send_uat"
    json_payload: dict = {
        "to": ["mark.tan@hanetf.com"],
        "subject": "Test subject",
        "cc": ["karen.kiwanuka@hanetf.com"],
        "bcc": ["ramon.williams@hanetf.com"],
        "body_text": "Test body",
        "body_html": None,
        "attachments": [r"Y:\dev-geoffrey-mark\geoffrey\tests\testdata\test_copy_files_to_sharepoint\Bar_list-RMAU-XS2115336336-all-all.xlsx"],
        "draft_or_send": "send"
    }

    row_data: dict = {"job_type": job_type, "json_payload": json_payload}
    print("sample_2 row_data: {}".format(json.dumps(row_data)))

def sample_3():
    table: str = "data.rundeck_job_queue"
    job_type: str = "graph_email_send_uat"

    html: str = """
            <h2 style="margin:0 0 8px 0;">HTML File Email Test</h2>

            <p>
              This is a <strong>bold</strong>, <em>italic</em>, and
              <span style="text-decoration:underline;">underlined</span> paragraph with a
              <a href="https://hanetf.com">link</a>.
            </p>

            <ul>
              <li>Bullet one</li>
              <li>Bullet two</li>
              <li>Bullet three</li>
            </ul>

            <table cellpadding="6" cellspacing="0" border="1" style="border-collapse:collapse;">
              <tr><th>Col A</th><th>Col B</th></tr>
              <tr><td>Alpha</td><td>1</td></tr>
              <tr><td>Beta</td><td>2</td></tr>
            </table>

            <p style="color:#666; margin-top:12px;">
              Footer note: line break test<br/>Second line here.
    """

    filepath_1: str = r"Y:\dev-geoffrey-mark\geoffrey\tests\testdata\test_copy_files_to_sharepoint\Bar_list-RMAU-XS2115336336-all-all.xlsx"

    json_payload: dict = {
        "to": "mark.tan@hanetf.com",
        "subject": "Test subject",
        "cc": ["karen.kiwanuka@hanetf.com"],
        "bcc": ["ramon.williams@hanetf.com"],
        "body_text": None,
        "body_html": html,
        "attachments": [filepath_1],
        "draft_or_send": "send"
    }

    row_data: dict = {"job_type": job_type, "json_payload": json_payload}
    print("sample_3 row_data: {}".format(json.dumps(row_data)))

def sample_4():
    table: str = "data.rundeck_job_queue"
    job_type: str = "graph_email_send_uat"

    html: str = """
                <h2 style="margin:0 0 8px 0;">HTML File Email Test</h2>
                <p>
                  This is a <strong>bold</strong>, <em>italic</em>, and
                  <span style="text-decoration:underline;">underlined</span> paragraph with a
                  <a href="https://hanetf.com">link</a>.
                </p>
            """

    filepath_1: str = r"Y:\dev-geoffrey-mark\geoffrey\tests\testdata\test_copy_files_to_sharepoint\Bar_list-RMAU-XS2115336336-all-all.xlsx"

    json_payload: dict = {
        "to": "mark.tan@hanetf.com",
        "subject": "Test subject",
        "cc": ["karen.kiwanuka@hanetf.com"],
        "bcc": ["ramon.williams@hanetf.com"],
        "body_text": None,
        "body_html": html,
        "attachments": [filepath_1],
        "draft_or_send": "send"
    }
    retry_max_attempts: int = 3  # optional. default=1
    retry_delay_minutes: int = 1  # optional. default=5
    context: dict = {"process": "sample testing"}  # optional. free form json

    row_data: dict = {"job_type": job_type,
                      "json_payload": json_payload,
                      "retry_max_attempts": retry_max_attempts,
                      "retry_delay_minutes": retry_delay_minutes,
                      "context": context}

    print("sample_4 row_data: {}".format(json.dumps(row_data)))

def sample_5():
    table: str = "data.rundeck_job_queue"
    job_type: str = "graph_email_send_uat"

    html: str = """
                <h2 style="margin:0 0 8px 0;">Sample 5 - Full Featured Email Test</h2>
                <p>
                  This email demonstrates all features with attachment, CC, and BCC.
                </p>
            """

    filepath_1: str = r"Y:\dev-geoffrey-mark\geoffrey\tests\testdata\test_copy_files_to_sharepoint\Bar_list-RMAU-XS2115336336-all-all.xlsx"

    json_payload: dict = {
        "to": "mark.tan@hanetf.com",
        "subject": "Test subject",
        "cc": ["karen.kiwanuka@hanetf.com"],
        "bcc": ["ramon.williams@hanetf.com"],
        "body_text": "Test body - plain text fallback",
        "body_html": html,
        "attachments": [filepath_1],
        "draft_or_send": "send"
    }
    retry_max_attempts: int = 3
    retry_delay_minutes: int = 1
    context: dict = {"process": "sample 5 testing"}

    row_data: dict = {"job_type": job_type,
                      "json_payload": json_payload,
                      "retry_max_attempts": retry_max_attempts,
                      "retry_delay_minutes": retry_delay_minutes,
                      "context": context}

    print("sample_5 row_data: {}".format(json.dumps(row_data)))


if __name__ == "__main__":
   sample_1()
   sample_2()
   sample_3()
   sample_4()
   sample_5()

