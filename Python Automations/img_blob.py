import base64 #need to install (pip install pybase64)

#converts an image to blob
def image_to_blob(image):
    with open(image,'rb') as image_file:
        blobData = base64.b64encode(image_file.read())
    return blobData

#converts an blob to image
def blob_to_image():
    img = open("image_saved.jpg","wb") #name the image going be saved
    blobData=image_to_blob("image_to_convert.jpg") #calling function to convert image to blob
    img.write(base64.b64decode(blobData)) #coverts a blob to image using b64decode
    img.close()

def main():
    blob_to_image()

if __name__ == "__main__":
    main()