from pdf2image import convert_from_bytes
import sys
from PIL import Image
from email.parser import BytesParser
from tnefparse.tnef import TNEF, TNEFAttachment, TNEFObject
import io

def make_thumb(pil_image, orig_filename):
  pil_image.thumbnail((180,180))
  pil_image.save('thumb_' + orig_filename)

def process_pdf(bytes, filename):
  pil_images = convert_from_bytes(bytes)
  basename = filename[0:-4] if filename.endswith('.pdf') else filename
  for i in range(len(pil_images)):
    jpgfilename = basename + '_' + str(i) + '.jpg'
    pil_images[i].save(jpgfilename, 'JPEG')
    make_thumb(pil_images[i], jpgfilename)

def process_image(bytes, filename):
  with open(filename, 'wb') as imagefile:
    imagefile.write(bytes)
  pil_image = Image.open(io.BytesIO(bytes))
  make_thumb(pil_image, filename)

f = open(sys.argv[1], 'rb')
msg = BytesParser().parse(f)

for part in msg.walk():
  filename = part.get_filename()
  [maintype, subtype] = part.get_content_type().split('/', 1)
  payload = part.get_payload(None, decode=True)
  if subtype == 'ms-tnef':
    tnef = TNEF(payload, do_checksum=True)
    for a in tnef.attachments:
      with open(a.name, "wb") as afp:
        afp.write(a.data)
  elif maintype == 'image':
    process_image(payload, filename)
  elif subtype == 'pdf':
    process_pdf(payload, filename)

f.close()
