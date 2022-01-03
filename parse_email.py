from pdf2image import convert_from_bytes
import sys
from PIL import Image
from email.parser import BytesParser
from tnefparse.tnef import TNEF, TNEFAttachment, TNEFObject
import io
import os

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

def process_tnef(bytes):
  tnef = TNEF(bytes, do_checksum=True)
  for a in tnef.attachments:
    content_type = None
    for attr in a.mapi_attrs:
      if attr.attr_type == 31 and attr.name == 14094:   # stupid microsoft enums
        content_type = attr.raw_data[0].rstrip('\x00')  # stupid microsoft nulls
    process_attachment(a.name, a.data, content_type)

def process_attachment(filename, payload, content_type):
  if not content_type:
    ext = os.path.splitext(filename)[1]
    content_type = {
      '.pdf': 'application/pdf'
    }.get(ext)
  if content_type:
    if content_type == 'application/ms-tnef':
      process_tnef(payload)
    elif content_type.startswith('image/'):
      process_image(payload, filename)
    elif content_type == 'application/pdf':
      process_pdf(payload, filename)

f = open(sys.argv[1], 'rb')
msg = BytesParser().parse(f)

for part in msg.walk():
  filename = part.get_filename()
  content_type = part.get_content_type()
  payload = part.get_payload(None, decode=True)
  process_attachment(filename, payload, content_type)

f.close()
