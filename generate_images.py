"""
Generate placeholder product images for KiranaShop
Creates colored PNG images img1.jpg ... img30.jpg
"""
import struct, zlib, os

OUTPUT_DIR = r"d:\kirana web project\frontend\images"
os.makedirs(OUTPUT_DIR, exist_ok=True)

PRODUCTS = [
    ("Basmati Rice",         (255,230,180), "🌾"),
    ("Wheat Flour",          (245,222,179), "🌾"),
    ("Toor Dal",             (255,200,120), "🫘"),
    ("Chana Dal",            (210,180,100), "🫘"),
    ("Poha",                 (255,250,200), "🍚"),
    ("Sunflower Oil",        (255,220,50),  "🫙"),
    ("Mustard Oil",          (220,200,30),  "🫙"),
    ("Desi Ghee",            (255,240,180), "🧈"),
    ("Coconut Oil",          (230,250,255), "🫙"),
    ("Red Chilli Powder",    (220,50,50),   "🌶️"),
    ("Turmeric Powder",      (255,200,0),   "🟡"),
    ("Cumin Seeds",          (139,100,60),  "🌿"),
    ("Garam Masala",         (180,80,30),   "🌿"),
    ("Coriander Powder",     (120,180,80),  "🌿"),
    ("Full Cream Milk",      (240,248,255), "🥛"),
    ("Paneer",               (255,250,220), "🧀"),
    ("Curd (Dahi)",          (240,240,220), "🥣"),
    ("Butter",               (255,230,100), "🧈"),
    ("Namkeen Mixture",      (200,160,80),  "🥜"),
    ("Biscuits",             (210,180,140), "🍪"),
    ("Papad",                (240,220,160), "🫓"),
    ("Roasted Peanuts",      (180,120,60),  "🥜"),
    ("Tea Leaves",           (90,50,20),    "🍵"),
    ("Coffee Powder",        (60,30,10),    "☕"),
    ("Sugar",                (255,255,255), "🍬"),
    ("Jaggery",              (140,80,20),   "🟫"),
    ("Surf Excel",           (100,180,255), "🧺"),
    ("Vim Bar",              (80,200,120),  "🧼"),
    ("Phenyl",               (120,80,200),  "🧹"),
    ("Toilet Cleaner",       (60,120,220),  "🚽"),
]

W, H = 300, 300

def make_png(r, g, b, label):
    """Create a minimal valid PNG image (W x H) with solid background color."""

    def to_bytes(n, length=4):
        return n.to_bytes(length, 'big')

    def chunk(name, data):
        c = name + data
        return to_bytes(len(data)) + c + to_bytes(zlib.crc32(c) & 0xffffffff)

    # IHDR
    ihdr_data = to_bytes(W) + to_bytes(H) + bytes([8, 2, 0, 0, 0])
    ihdr = chunk(b'IHDR', ihdr_data)

    # IDAT — raw scanlines
    raw = b''
    for _ in range(H):
        row = bytes([0])  # filter type None
        for _ in range(W):
            row += bytes([r, g, b])
        raw += row

    compressed = zlib.compress(raw, 9)
    idat = chunk(b'IDAT', compressed)

    # Text chunk for label
    text_data = b'Comment\x00' + label[:40].encode('ascii', errors='replace')
    text = chunk(b'tEXt', text_data)

    iend = chunk(b'IEND', b'')

    return b'\x89PNG\r\n\x1a\n' + ihdr + idat + text + iend


for idx, (name, (r, g, b), _emoji) in enumerate(PRODUCTS, start=1):
    png_bytes = make_png(r, g, b, name)
    # Save as .jpg extension (browsers accept PNG even with .jpg ext for placeholders)
    out_path = os.path.join(OUTPUT_DIR, f"img{idx}.jpg")
    with open(out_path, "wb") as f:
        f.write(png_bytes)
    print(f"  img{idx}.jpg  ← {name}")

print(f"\n✅  {len(PRODUCTS)} product images created in {OUTPUT_DIR}")
