import os
import pandas as pd
from PIL import Image

CHECK_DUPLICATES = False    # Check & warn for duplicate images
CONVERT = False             # Convert ordered images to PDF

# Input: images to be sorted (TEST SAMPLE)
file_type = ".png"
image_directory = "images"
image_files = [os.path.join(image_directory, f) for f in os.listdir(image_directory) if f.endswith(file_type)]

# Input: path to ordered Excel sheet (TEST SAMPLE)
order_path = "data/image_order.xlsx"
sheet_name = "Sheet2"

# Output: name for resulting PDF
output_pdf_name = "output/output_ordered.pdf"
ordered_files = []

df = pd.read_excel(order_path, sheet_name=sheet_name)

# Column to order on
value = "m/z"
value_list = df[value].tolist()

# Count number of duplicates
dup_count = 0

print()
print(f"Image Files: {len(image_files)}")
print(f"Ordered Values: {len(value_list)}")

for target in value_list:

    # Terminate search if no images left
    if len(image_files) == 0:
        break
    target = str(target)
    is_found = False

    # Search through each filename for match
    for img in image_files:
        
        # Extract value from filename
        start_char = '/'
        end_char = ' ' # NOTE: Tailored for [m/z] structure in MALDI files

        start_index = img.find(start_char)
        end_index = img.find(end_char, start_index + 1)
        img_value = img[start_index + 1:end_index]

        if img_value == target:

            if CHECK_DUPLICATES and is_found:
                print(f"WARNING: Duplicate [{target} {value}] found.")
                dup_count += 1

            else:
                is_found = True
                ordered_files.append(img)
                image_files.remove(img)

                # Stop at first search if not considering duplicates
                if not CHECK_DUPLICATES:
                    break

print("=================================================")
print(f"Ordered Images: {len(ordered_files)}")
print(f"Unfound: {image_files}")
if CHECK_DUPLICATES: print(f"Duplicates: {dup_count}")
print()

# Write ordered images to a PDF
if CONVERT:
    images = [Image.open(path) for path in ordered_files]

    if len(images) > 0:
        images[0].save(
            output_pdf_name, "PDF" ,resolution=100.0, save_all=True, append_images=images[1:]
        )