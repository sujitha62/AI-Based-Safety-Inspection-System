"""
Training script for the Building Defect Detection model (YOLOv8).

IMPORTANT: the previous version of this script used:
    epochs=1, fraction=0.001, imgsz=320, val=False
That trains on ~14 images for 1 epoch -> the model never actually learns
anything, which is why it couldn't detect cracks. This version trains on
the FULL dataset with a real epoch count.

Before running this:
1. Confirm your `labels/` folder contains YOLO-format .txt files
   (class_id x_center y_center width height, normalized 0-1),
   NOT the original VOC XML files. If you haven't converted yet,
   use voc_to_yolo.py first.
2. Confirm data.yaml paths point to the correct train/val image folders.
3. If you have a GPU, this will be much faster. On CPU-only, expect
   this to take a long time on 14k+ images — consider starting with
   epochs=30-50 to get a baseline result first, then increasing.
"""

from ultralytics import YOLO


def main():
    # yolov8n = fastest/smallest (good for CPU or quick iteration)
    # yolov8s = better accuracy, still reasonably fast
    model = YOLO("yolov8n.pt")

    model.train(
        data="data.yaml",
        epochs=100,            # real training run, not a smoke test
        imgsz=640,             # standard YOLO input size (better than 320)
        batch=16,              # lower to 8 if you run out of memory
        workers=4,             # set to 0 if you hit multiprocessing issues on Windows
        patience=20,           # stop early if no improvement for 20 epochs
        val=True,              # actually validate during training
        save=True,
        project="runs/detect",
        name="train_full",     # new run name so it doesn't overwrite train-4
        # Light augmentation helps a lot on a defect dataset like this:
        degrees=5,
        fliplr=0.5,
        mosaic=1.0,
    )


if __name__ == "__main__":
    # The __main__ guard is required on Windows when workers > 0
    main()