import os
import aspose.words as aw


def list_of_files(filepath):
    filenames = []
    extensions = (".tif", ".TIF")
    for subdirs, dir, files in os.walk(filepath):
        for file in files:
            file_path = os.path.join(subdirs, file)
            if file.endswith(extensions):
                filenames.append(file_path)
    return filenames


filepath = input("Enter folder path: ").replace('"', '').replace("'", "")
doc = aw.Document()
builder = aw.DocumentBuilder(doc)
files = list_of_files(filepath)
if len(files) == 0:
    print("No TIF files in the supplied directory.")
else:
    for file in files:
        img = builder.insert_image(file)
        filename = os.path.basename(file).split(".")[0]
        dirname = os.path.dirname(file)
        if not os.path.exists(f"{dirname}/PNGS/"):
            os.mkdir(f"{dirname}/PNGs/")
        img.image_data.save(f"{dirname}/PNGs/{filename}.png")
        print(f"{filename} saved to {dirname}/PNGs/")
