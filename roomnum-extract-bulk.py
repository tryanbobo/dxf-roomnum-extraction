import ezdxf
import os
import pandas as pd


def extract_text_from_dxf(dxf_file):
    try:
        text = ""
        dwg = ezdxf.readfile(dxf_file)
        for entity in dwg.entities:
            if entity.dxftype() == 'TEXT':
                text += entity.dxf.text + '\n'
            elif entity.dxftype() == 'MTEXT':
                text += entity.plain_text() + '\n' # get the MTEXT too.
        return text
    except Exception as e:
        print("Error extracting text from {}: {}".format(dxf_file, e))
        return ""


def extract_room_numbers(text):
    room_numbers = set()  # Using a set to store unique values
    lines = text.split('\n')
    excluded_substrings = ["user", "men", "women", "cust", "closet", "plate", "up", "down", "above", "below"]
    for line in lines:
        line = line.strip()
        if 3 <= len(line) < 7 and not any(substring in line.lower() for substring in excluded_substrings): # Skip if 'x' in room name
            room_numbers.add(line)
    return sorted(room_numbers)  # Sort and return unique values


def determine_floor(room_numbers):
    counts = {}
    for room in room_numbers:
        if room[0].isdigit():
            counts[room[0]] = counts.get(room[0], 0) + 1

    if counts:
        floor = max(counts, key=counts.get)
    else:
        floor = ""

    return floor


def process_dxf(dxf_file):
    try:
        extracted_text = extract_text_from_dxf(dxf_file)
        if extracted_text:
            room_numbers = extract_room_numbers(extracted_text)

            # Skip if no room numbers starting with digits
            #if not any(room[0].isdigit() for room in room_numbers):
            #    print("No room numbers starting with digits in:", dxf_file)
            #    return

            floor = determine_floor(room_numbers)
            data = {
                'Building': os.path.basename(dxf_file),
                'Floor': floor,
                'Room': room_numbers
            }
            df = pd.DataFrame(data)
            output_file = os.path.splitext(dxf_file)[0] + "_room-numbers.xlsx"
            df.to_excel(output_file, index=False)
            print("Excel file created:", output_file)
        else:
            print("No text extracted or error occurred for file:", dxf_file)
    except Exception as e:
        print("Error processing {}: {}".format(dxf_file, e))


# Example usage
root_folder = r"M:\MapCom\Bldg_floorplans\_DxfOutput"

for root, dirs, files in os.walk(root_folder):
    for filename in files:
        if filename.endswith(".dxf"):
            dxf_file = os.path.join(root, filename)
            process_dxf(dxf_file)