import emoji
import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

# Define status mapping based on sprite sheet positions
STATUS_MAPPING = {
    "0px -399px": "Not Boosted",
    "0px -105px": "Auto-Generated Captions and Subtitles",
    "0px -21px": "Not a Crosspost",
    "0px -42px": "Couldn’t Auto-Generate Captions",
    "0px 0px": "Crosspost",
}

# Path to ChromeDriver
CHROMEDRIVER_PATH = "/Users/maxwellthomason/Downloads/chromedriver-mac-arm64/chromedriver"

# Directory to save the output file
OUTPUT_DIRECTORY = "/Users/maxwellthomason/Downloads/chromedriver-extracted/chromedriver-mac-arm64"

# Initialize WebDriver
service = Service(CHROMEDRIVER_PATH)
driver = webdriver.Chrome(service=service)

# Open the Facebook Business Suite URL
url = "https://business.facebook.com/latest/posts/published_posts?business_id=587093741487437&asset_id=111803458179936&should_show_nux=false&focus_comments=false"
driver.get(url)

# Wait for user to press Enter
input("Please log in to Facebook and complete any required actions. Once you're ready, press Enter to continue...")

# Function to select "Lifetime" in the time range dropdown
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def select_lifetime_time_range():
    try:
        # Wait for the dropdown to be visible
        time_range_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//div[contains(text(), "Last 90 days")]'))
        )
        time_range_button.click()

        # Wait for lifetime option to appear
        lifetime_option = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//input[@value="LIFETIME"]'))
        )
        lifetime_option.click()
    except Exception as e:
        print(f"Error setting time range to 'Lifetime': {e}")

# Function to configure columns
def configure_columns(categories):
    try:
        columns_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//div[contains(text(), "Columns")]'))
        )
        columns_button.click()
        
        for category in categories:
            try:
                category_button = driver.find_element(By.XPATH, f'//div[contains(text(), "{category}")]/ancestor::div[contains(@class, "x1n2onr6")]')
                checkbox_input = category_button.find_element(By.XPATH, './/input[@type="checkbox"]')
                if checkbox_input.get_attribute("aria-checked") == "true":
                    print(f"Category '{category}' already selected. Skipping...")
                    continue
                print(f"Selecting category '{category}'...")
                driver.execute_script("arguments[0].click();", category_button)
                time.sleep(1)  # Small delay
            except Exception as e:
                print(f"Error interacting with category '{category}': {e}")
        
        apply_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//div[contains(text(), "Apply")]'))
        )
        apply_button.click()
    except Exception as e:
        print(f"Error configuring columns: {e}")

# Function to scroll and load posts
def scroll_to_load_posts():
    SCROLL_PAUSE_TIME = 3
    MAX_SCROLLS = 500
    RETRY_LIMIT = 3
    scroll_count = 0
    retry_count = 0
    prev_row_count = 0
    while scroll_count < MAX_SCROLLS:
        driver.execute_script("window.scrollBy(0, document.body.scrollHeight);")
        time.sleep(SCROLL_PAUSE_TIME)
        
        # Wait for rows to load
        rows = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "tbody > tr"))
        )
        current_row_count = len(rows)
        if current_row_count == prev_row_count:
            retry_count += 1
            if retry_count >= RETRY_LIMIT:
                print(f"Stopped scrolling after {scroll_count} attempts. Rows loaded: {prev_row_count}")
                break
        else:
            retry_count = 0
        prev_row_count = current_row_count
        scroll_count += 1
        print(f"Scroll count: {scroll_count} - Rows currently loaded: {current_row_count}")

# Initialize the data dictionary
data = {"Title": []}  # Initialize with Title key

def extract_title_metrics(columns):
    """
    Extract title-related metrics including emoji count, word count, and character count.
    """
    try:
        # Locate the title element using the updated CSS selector
        try:
            title_element = columns[1].find_element(By.CSS_SELECTOR, 
                ".xmi5d70.x1fvot60.xo1l8bm.xxio538.xbsr9hj.xuxw1ft.x6ikm8r.x10wlt62.xlyipyv.x1h4wwuj.xeuugli.x1uvtmcs.xh8yej3#js_td")
            title_text = title_element.text.strip()
        except Exception:
            title_text = "Unknown"
        
        # Calculate metrics
        from emoji import is_emoji

        emoji_count = len([char for char in title_text if is_emoji(char)])
        word_count = len(title_text.split())
        char_count = len(title_text)

        # Return metrics
        return {
            "Emoji Count": emoji_count,
            "Word Count": word_count,
            "Character Count": char_count,
        }

    except Exception as e:
        print(f"Error extracting title metrics: {e}")
        return {
            "Emoji Count": 0,
            "Word Count": 0,
            "Character Count": 0,
        }

def extract_statuses(columns):
    """
    Extract all possible statuses for a post based on the sprite sheet crop locations.
    :param columns: List of <td> elements in the current row.
    :return: Dictionary with status columns and binary indicators (1 or 0).
    """
    status_dict = {status: 0 for status in STATUS_MAPPING.values()}  # Initialize all statuses to 0

    try:
        # Locate the status <td> element (6th column)
        status_cell = columns[5]  # aria-colindex starts at 6 (index 5 for zero-based index)

        # Find all <span> elements with class xsgj6o6 within the cell
        span_elements = status_cell.find_elements(By.CSS_SELECTOR, "span.xsgj6o6")
        for span in span_elements:
            try:
                icon_element = span.find_element(By.CSS_SELECTOR, "i.img")
                style = icon_element.get_attribute("style")
                print(f"Extracted style: {style}")  # Debugging line

                # Extract the background-position value
                position = None
                if "background-position" in style:
                    position = style.split("background-position:")[1].split(";")[0].strip()
                    print(f"Extracted position: {position}")  # Debugging line

                # If the position matches a known status, mark it as present
                if position in STATUS_MAPPING:
                    status = STATUS_MAPPING[position]
                    print(f"Matched status: {status}")  # Debugging line
                    status_dict[status] = 1  # Set the status to 1
                else:
                    print(f"No match for position: {position}")  # Debugging line
            except Exception as e:
                print(f"Error processing status span: {e}")
                continue
    except Exception as e:
        print(f"Error extracting statuses: {e}")

    return status_dict

# Function to extract Post Type, Post Location, and Title Metrics
def extract_post_type_location_title(rows):
    """
    Extract post type, location, and title metrics for all rows.
    """
    print("Extracting Post Type, Post Location, and Title Metrics...")
    post_data = {
        "Post Type": [],
        "Post Location": [],
    }

    for i, row in enumerate(rows):
        try:
            columns = row.find_elements(By.TAG_NAME, "td")

            # Extract Post Type
            try:
                post_type_element = columns[1].find_element(By.CSS_SELECTOR, "div")
                post_type_text = post_type_element.text.strip().lower()
                if "reel" in post_type_text:
                    post_data["Post Type"].append("Reel")
                elif "photo" in post_type_text:
                    post_data["Post Type"].append("Photo")
                elif "video" in post_type_text:
                    post_data["Post Type"].append("Video")
                elif "text" in post_type_text:
                    post_data["Post Type"].append("Text")
                else:
                    post_data["Post Type"].append("Link")
            except Exception:
                post_data["Post Type"].append(pd.NA)

            # Extract Post Location
            try:
                img_elements = columns[1].find_elements(By.TAG_NAME, "img")
                location_found = False
                for img in img_elements:
                    alt_text = img.get_attribute("alt")
                    if alt_text in ["Facebook", "Instagram"]:
                        post_data["Post Location"].append(alt_text)
                        location_found = True
                        break
                if not location_found:
                    post_data["Post Location"].append(pd.NA)
            except Exception:
                post_data["Post Location"].append(pd.NA)

            # Extract Statuses
            statuses = extract_statuses(columns)
            for status, value in statuses.items():
                if status not in post_data:
                    post_data[status] = []
                post_data[status].append(value)
            
        except Exception as e:
            print(f"Error processing row {i + 1}: {e}")
            post_data["Post Type"].append(pd.NA)
            post_data["Post Location"].append(pd.NA)

    return post_data

# Title Calculated Values
def calculate_title_metrics(df):
    from emoji import is_emoji  # Ensure the correct emoji method is imported
    df["Emoji Count"] = df["Title"].apply(lambda title: len([char for char in title if is_emoji(char)]) if pd.notna(title) else 0)
    df["Word Count"] = df["Title"].apply(lambda title: len(title.split()) if pd.notna(title) else 0)
    df["Character Count"] = df["Title"].apply(lambda title: len(title) if pd.notna(title) else 0)
    return df

# Function to extract headers
def extract_headers():
    headers = {}
    try:
        header_elements = driver.find_elements(By.CSS_SELECTOR, "thead th")
        for idx, header in enumerate(header_elements, start=1):
            header_name = header.text.strip() if header.text.strip() else f"Unknown_{idx}"
            headers[idx - 1] = header_name

        # Ensure "Title" column is explicitly added
        if "Title" not in headers.values():
            headers[len(headers)] = "Title"
    except Exception as e:
        print(f"Error extracting headers: {e}")
    return headers

def extract_data(headers):
    """
    Extract data rows, including dynamically scraped headers.
    """
    # Ensure "Title" exists in headers
    if "Title" not in headers.values():
        headers[len(headers)] = "Title"

    # Initialize data dictionary with headers
    data = {header_name: [] for header_name in headers.values()}

    for i, row in enumerate(rows):
        print(f"Processing row {i + 1} of {len(rows)}...")

        try:
            # Get all <td> elements in the row
            columns = row.find_elements(By.TAG_NAME, "td")

            # Handle missing or incomplete rows
            if len(columns) < 2:  # If fewer than 2 columns, skip and pad with pd.NA
                print(f"Row {i + 1} has fewer than 2 columns. Skipping.")
                for header_name in headers.values():
                    data[header_name].append(pd.NA)
                continue

            # Extract Title
            try:
                # Locate the title element using the CSS selector
                title_element = columns[1].find_element(By.CSS_SELECTOR, "div#js_tb")
                title = title_element.text.strip()
            except Exception as e:
                print(f"Primary selector failed for Title in row {i + 1}. Error: {e}")
                title = columns[1].text.strip() if len(columns) > 1 and columns[1].text else "Unknown"

            # Append the title to the data
            data["Title"].append(title if title else pd.NA)

            # Extract remaining columns dynamically
            for idx, header_name in headers.items():
                if header_name == "Title":
                    continue  # Skip "Title" since it's already handled
                try:
                    if idx < len(columns):
                        cell = columns[idx].find_element(By.CSS_SELECTOR, "span") if idx < len(columns) else None
                        cell_value = cell.text.strip() if cell else pd.NA
                        data[header_name].append(cell_value)
                    else:
                        print(f"Missing column for '{header_name}' in row {i + 1}")
                        data[header_name].append(pd.NA)
                except Exception as e:
                    print(f"Error extracting '{header_name}' for row {i + 1}: {e}")
                    data[header_name].append(pd.NA)

        except Exception as e:
            print(f"Error processing row {i + 1}: {e}")
            # Append pd.NA for all headers in case of row-level failure
            for header_name in headers.values():
                data[header_name].append(pd.NA)

    return data

# Extract Date Published
def extract_date_and_time(rows):
    """
    Extract 'Date Published' and split into separate Date and Time columns.
    Handles cases where time is not available or the column structure is missing.
    """
    date_list = []
    time_list = []

    for i, row in enumerate(rows):
        try:
            print(f"Processing 'Date Published' for row {i + 1}...")

            # Safely locate the 5th column for 'Date Published'
            columns = row.find_elements(By.TAG_NAME, "td")
            if len(columns) < 5:
                raise IndexError(f"Row {i + 1} does not have a 'Date Published' column.")

            # Extract the 'Date Published' text
            date_cell = columns[4]  # 5th column
            date_div = date_cell.find_element(
                By.CSS_SELECTOR,
                "div.xmi5d70.x1fvot60.xo1l8bm.xxio538.xbsr9hj.xuxw1ft.x6ikm8r.x10wlt62.xlyipyv.x1h4wwuj.xeuugli.x1uvtmcs"
            )
            date_published = date_div.text.strip()

            # Split into date and time if time is included
            if date_published.endswith("am") or date_published.endswith("pm"):
                date, time = date_published.rsplit(",", 1)
                date_list.append(date.strip())
                time_list.append(time.strip())
            else:
                date_list.append(date_published.strip())
                time_list.append("")  # No time provided
        except Exception as e:
            print(f"Error processing 'Date Published' for row {i + 1}: {e}")
            date_list.append(None)
            time_list.append(None)

    return date_list, time_list

# Select Lifetime Time Range and Configure Columns
select_lifetime_time_range()
configure_columns(["Content details", "Performance", "Engagement", "Video", "Paid"])

# Scroll and Load Posts
scroll_to_load_posts()

rows = driver.find_elements(By.CSS_SELECTOR, "tbody > tr")
date_list, time_list = extract_date_and_time(rows)

if not rows:
    print("No rows found. Exiting script.")
    driver.quit()
    exit(1)

# Extract Headers and Data
headers = extract_headers()
print(f"Extracted Headers: {headers}")

# Safeguard against missing rows
if not rows:
    print("No rows found. Exiting script.")
    driver.quit()
    exit(1)

# Extract visible data
data = extract_data(headers)

# Ensure missing keys exist in post_data
post_data = extract_post_type_location_title(rows)

# Handle missing keys in post_data and update the data dictionary
missing_keys = set(post_data.keys()) - set(data.keys())
for key in missing_keys:
    data[key] = [pd.NA] * len(data[list(data.keys())[0]])
data.update(post_data)

# Add Date and Time to the data dictionary
data["Date"] = date_list
data["Time"] = time_list

# Verify column lengths
max_len = max(len(col) for col in data.values())
for header, col in data.items():
    if len(col) != max_len:
        print(f"Column '{header}' length mismatch: {len(col)} (expected {max_len}). Padding with NaN.")
        while len(col) < max_len:
            col.append(pd.NA)

# Create DataFrame and remove columns with 'Unknown' as a header
df = pd.DataFrame(data).loc[:, ~pd.DataFrame(data).columns.str.contains('^Unknown', case=False)]

# Remove Date Published Column (split to date/time)
if "Date published" in df.columns:
    df.drop(columns=["Date published"], inplace=True)
    print("Removed 'Date published' column from the DataFrame.")
else:
    print("'Date published' column not found in the DataFrame. Skipping removal.")

# Calculate Title Metrics
df = calculate_title_metrics(df)

# Convert specific columns to numeric in the DataFrame
numeric_columns = [
    "Reach",
    "Likes and reactions",
    "Shares",
    "Comments",
    "Impressions",
    "Plays",
    "Saves",
    "Interactions",
    "Watch time",
    "Average watch time",
    "Link clicks",
    "Approximate in-stream ad earnings",
]

for column in numeric_columns:
    if column in df.columns:
        if column in ["Watch time", "Average watch time", "Link clicks", "Approximate in-stream ad earnings"]:
            # Convert the column to string before replacing values
            df[column] = df[column].astype(str).replace("--", "")
            # Convert "HH:MM" format to seconds for "Average watch time"
            if column == "Average watch time":
                df[column] = df[column].str.split(":").apply(
                    lambda x: int(x[0]) * 60 + int(x[1]) if len(x) == 2 else pd.NA
                )
            # Convert to numeric
            df[column] = pd.to_numeric(df[column], errors="coerce")
        else:
            # General numeric conversion for other columns
            df[column] = pd.to_numeric(df[column], errors="coerce")
    else:
        print(f"Warning: Column '{column}' not found. Skipping numeric conversion.")

# Define the function to add calculated fields
def add_calculated_fields(df):
    print("Adding calculated fields...")

    # Engagement Rate (%)
    try:
        df["Engagement Rate (%)"] = df.apply(
            lambda row: round(((row["Likes and reactions"] + row["Comments"] + row["Shares"]) / max(row["Reach"], 1)) * 100, 2)
            if pd.notna(row["Reach"]) and row["Reach"] > 0 else "",
            axis=1
        )
    except KeyError as e:
        print(f"Missing column required for Engagement Rate: {e}")

    # Weighted Engagement Score
    df["Weighted Engagement Score"] = df.apply(
        lambda row: row["Likes and reactions"] + (2 * row["Comments"]) + (3 * row["Shares"])
        if pd.notna(row["Likes and reactions"]) and row["Likes and reactions"] > 0 else pd.NA, axis=1
    )

    # Comments-to-Likes Ratio
    df["Comments-to-Likes Ratio"] = df.apply(
        lambda row: round(row["Comments"] / max(row["Likes and reactions"], 1), 2)
        if pd.notna(row["Likes and reactions"]) and row["Likes and reactions"] > 0 else pd.NA, axis=1
    )

    # Shares-to-Likes Ratio
    df["Shares-to-Likes Ratio"] = df.apply(
        lambda row: round(row["Shares"] / max(row["Likes and reactions"], 1), 2)
        if pd.notna(row["Likes and reactions"]) and row["Likes and reactions"] > 0 else pd.NA, axis=1
    )

    # Interaction-to-Impressions Ratio (%)
    if "Impressions" in df.columns:
        df["Interaction-to-Impressions Ratio (%)"] = df.apply(
            lambda row: round(((row["Likes and reactions"] + row["Comments"] + row["Shares"]) / max(row["Impressions"], 1)) * 100, 2)
            if pd.notna(row["Impressions"]) and row["Impressions"] > 0 else pd.NA, axis=1
        )

    return df


# Ensure required columns are present and fill with NaN if missing
required_columns = ["Reach", "Likes and reactions", "Shares", "Comments", "Impressions", "Saves", "Interactions", "Watch time"]
for column in required_columns:
    if column not in df.columns:
        print(f"Warning: Missing column '{column}'. Filling with default values (null).")
        df[column] = pd.NA

# Add calculated fields
df = add_calculated_fields(df)

# Replace "--" and "Not enough data" with blank strings
df.replace(["--", "Not enough data"], "", inplace=True)

# Ensure the output directory exists
if not os.path.exists(OUTPUT_DIRECTORY):
    try:
        os.makedirs(OUTPUT_DIRECTORY)
        print(f"Created output directory: {OUTPUT_DIRECTORY}")
    except Exception as e:
        print(f"Error creating output directory: {e}")
        exit(1)
    print(f"Created output directory: {OUTPUT_DIRECTORY}")

# Create the output file path
from datetime import datetime
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
output_file = os.path.join(OUTPUT_DIRECTORY, f"facebook_posts_{timestamp}.xlsx")

# Save the DataFrame to Excel
print(f"Attempting to save Excel file to: {output_file}")
try:
    df.to_excel(output_file, index=False)
    print(f"Data scraping completed. Saved to '{output_file}'.")
except Exception as e:
    print(f"Error saving Excel file: {e}")
