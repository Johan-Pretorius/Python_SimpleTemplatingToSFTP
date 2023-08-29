import os
import openpyxl
from string import Template
import shutil
import pysftp


def clear_output_folder(output_folder):
    if os.path.exists(output_folder):
        shutil.rmtree(output_folder)
    os.makedirs(output_folder)


def upload_to_sftp(host, port, username, password, local_path, remote_path):
    cnopts = pysftp.CnOpts()
    cnopts.hostkeys = None  # Disable host key checking

    with pysftp.Connection(
        host, port=port, username=username, password=password, cnopts=cnopts
    ) as sftp:
        # Upload files to the remote directory
        for output_file in os.listdir(local_path):
            local_file_path = os.path.join(local_path, output_file)
            remote_file_path = os.path.join(remote_path, output_file)

            # Check if local file exists before uploading
            if os.path.exists(local_file_path):
                print("YAHOOOOOOOO!!!!!!!!!!!")
                sftp.put(local_file_path, remote_file_path)

            else:
                print(f"Local file not found: {local_file_path}")

        print("Upload completed successfully.")


# Function to read Excel data
def read_excel_data(file_path):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    data = []
    for row in sheet.iter_rows(values_only=True):
        data.append(row)
    return data


# Function to process template and generate output messages
def process_template(template_path, data, output_folder):
    with open(template_path, "r") as template_file:
        template_content = template_file.read()
        template = Template(template_content)

        headers = data[0]  # First row contains headers
        print("Headers:", headers)
        print("Data:", data[1:])
        template_file_name = str.replace(template_file.name, ".txt", "")
        template_file_name = str.replace(template_file_name, "templates\\", "")
        print("template_file_name:", template_file_name)

        for index, row in enumerate(data[1:], start=1):  # Start from the second row
            data_dict = {headers[i]: cell for i, cell in enumerate(row)}
            formatted_message = template.safe_substitute(data_dict)

            output_file_path = os.path.join(
                output_folder, f"{template_file_name}_output_{index}.in"
            )

            with open(output_file_path, "w") as output_file:
                output_file.write(formatted_message)


# Function to display a paginated list of templates
def display_paginated_templates(template_choices):
    per_page = 20
    total_templates = len(template_choices)

    current_page = 0

    while current_page * per_page < total_templates:
        start_idx = current_page * per_page
        end_idx = start_idx + per_page
        templates_to_display = template_choices[start_idx:end_idx]
        page_number = current_page + 1
        print(f"Page {page_number}:\n")

        for i, template in enumerate(templates_to_display, start=start_idx):
            print(f"{i+1}. {template}")

        choice = input(
            "\nEnter the number of the template, 'n' for next page, 'p' for previous page, or 'q' to quit: "
        ).lower()

        if choice == "q":
            return None  # User chose to quit
        elif choice == "n":
            current_page += 1
        elif choice == "p":
            if current_page > 0:
                current_page -= 1
            else:
                print("You are already on the first page.")
        else:
            try:
                choice = int(choice)
            except ValueError:
                print("Invalid input. Please enter a number, 'n', 'p', or 'q'.")
                continue

            if choice >= 1 and choice <= len(template_choices):
                return template_choices[choice - 1]
            else:
                print("Invalid choice. Please enter a valid number.")

    return None  # No template selected


# Main function to orchestrate the process
def main():
    templates_folder = "templates"
    data_folder = "data"
    output_folder = "output"

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    template_files = [f for f in os.listdir(templates_folder) if f.endswith(".txt")]

    template_choices = [
        os.path.splitext(template_file)[0] for template_file in template_files
    ]
    selected_template_name = display_paginated_templates(template_choices)

    if selected_template_name is None:
        return  # User chose to quit

    selected_template_path = os.path.join(
        templates_folder, f"{selected_template_name}.txt"
    )
    selected_data_path = os.path.join(data_folder, f"{selected_template_name}.xlsx")

    if not os.path.exists(selected_data_path):
        print(f"Error: Data file {selected_template_name}.xlsx not found.")
        return

    data = read_excel_data(selected_data_path)
    #    data = read_excel_data(selected_data_file)

    clear_output_folder(output_folder)

    process_template(selected_template_path, data, output_folder)
    #    process_template(selected_template, data, output_folder)
    print("Messages have been generated and saved in the 'output' folder.")

    # Ask the user if they want to upload to FTP
    upload_to_ftp = input(
        "Do you want to upload the output files to the FTP server? (y/n): "
    )

    if upload_to_ftp.lower() == "y":
        ftp_host = "xxx.xxx.xxx.xxx"
        ftp_username = "user"
        ftp_password = "password"
        ftp_port = 22
        remote_path = (
            "/somePath/"
        )

        upload_to_sftp(
            host=ftp_host,
            username=ftp_username,
            password=ftp_password,
            port=ftp_port,
            local_path=output_folder,
            remote_path=remote_path,
        )
        print("Output files have been uploaded to the FTP server.")


if __name__ == "__main__":
    main()
