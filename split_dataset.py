"""
Splits the MBDD2025 dataset into separate train/ and val/ folders.

Expected structure:
    project/
      MBDD2025/
        images/
        labels/
        Annotations/

Produces:
    project/
      dataset/
        train/
          images/
          labels/
        val/
          images/
          labels/
"""

import os
import random
import shutil

# MBDD2025 dataset paths
IMAGES_DIR = "MBDD2025/images"
LABELS_DIR = "MBDD2025/labels"

# Output folder
OUTPUT_DIR = "dataset"

# 15% validation, 85% training
VAL_RATIO = 0.15

SEED = 42


def main():
    random.seed(SEED)

    # Check folders
    if not os.path.exists(IMAGES_DIR):
        print(f"ERROR: Images folder not found: {IMAGES_DIR}")
        return

    if not os.path.exists(LABELS_DIR):
        print(f"ERROR: Labels folder not found: {LABELS_DIR}")
        return

    # Find images
    image_files = [
        f for f in os.listdir(IMAGES_DIR)
        if f.lower().endswith((".jpg", ".jpeg", ".png"))
    ]

    print(f"Found {len(image_files)} images.")

    # Shuffle images
    random.shuffle(image_files)

    # Split
    val_count = int(len(image_files) * VAL_RATIO)

    val_files = image_files[:val_count]
    train_files = image_files[val_count:]

    print(f"Training images: {len(train_files)}")
    print(f"Validation images: {len(val_files)}")

    # Create train and validation folders
    for split_name, files in [
        ("train", train_files),
        ("val", val_files)
    ]:

        img_out = os.path.join(
            OUTPUT_DIR,
            split_name,
            "images"
        )

        lbl_out = os.path.join(
            OUTPUT_DIR,
            split_name,
            "labels"
        )

        os.makedirs(img_out, exist_ok=True)
        os.makedirs(lbl_out, exist_ok=True)

        copied = 0
        skipped = 0

        for img_name in files:

            base_name = os.path.splitext(img_name)[0]

            label_name = base_name + ".txt"

            src_img = os.path.join(
                IMAGES_DIR,
                img_name
            )

            src_lbl = os.path.join(
                LABELS_DIR,
                label_name
            )

            # Skip images without labels
            if not os.path.exists(src_lbl):
                skipped += 1
                continue

            # Copy image
            shutil.copy2(
                src_img,
                os.path.join(img_out, img_name)
            )

            # Copy label
            shutil.copy2(
                src_lbl,
                os.path.join(lbl_out, label_name)
            )

            copied += 1

        print(
            f"{split_name}: "
            f"copied {copied} image+label pairs, "
            f"skipped {skipped}"
        )

    print("\nDataset splitting completed successfully!")

    print("\nUse this data.yaml:")

    print("""
path: dataset

train: train/images
val: val/images

nc: 5

names:
  - crack
  - leakage
  - corrosion
  - abscission
  - bulge
""")


if __name__ == "__main__":
    main()