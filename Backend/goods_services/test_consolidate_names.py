import cv2, numpy as np, pytesseract, sys
img=cv2.imread(sys.argv[1]); H,W=img.shape[:2]
# 1 plain OCR of whole page region below the signature
plain=pytesseract.image_to_string(img[int(H*.7):,:W//2])
print("PLAIN:",repr(plain.strip()))
# 2 colour mask: stamp ink is blue/purple, table text is black
hsv=cv2.cvtColor(img,cv2.COLOR_BGR2HSV)
mask=cv2.inRange(hsv,(105,40,60),(150,255,255))
mask=cv2.morphologyEx(mask,cv2.MORPH_CLOSE,np.ones((5,5),np.uint8))
n,lab,st,_=cv2.connectedComponentsWithStats(mask)
# find biggest roughly-square blob = stamp
cands=[(s[4],s) for s in st[1:] if 0.7<s[2]/max(s[3],1)<1.4 and s[2]>150]
if not cands: print("no stamp found");sys.exit()
x,y,w,h,_=max(cands,key=lambda c:c[0])[1]
print("stamp box",x,y,w,h)
crop=mask[y:y+h,x:x+w]; cv2.imwrite('stamp_mask.png',255-crop)
# 3 polar unwrap -> straight text
cx,cy=w/2,h/2; R=min(w,h)/2
pol=cv2.warpPolar(crop,(int(R),int(2*np.pi*R)),(cx,cy),R,cv2.WARP_POLAR_LINEAR)
pol=cv2.rotate(pol,cv2.ROTATE_90_COUNTERCLOCKWISE)
band=pol[:, int(pol.shape[1]*0.0):]
band=255-cv2.resize(band,None,fx=2,fy=2,interpolation=cv2.INTER_CUBIC)
cv2.imwrite('unwrapped.png',band)
# repeat strip twice so text crossing the 0/360 seam is whole, try each 90deg rotation of the ring
for name,im in [('mask',255-crop)]:
    print(name,repr(pytesseract.image_to_string(im,config='--psm 11').strip()))
wide=np.hstack([band,band])
print("UNWRAP:",repr(pytesseract.image_to_string(wide,config='--psm 6').strip()))
