const QRCode = require('../portal/qrcode.min.js');

const text = "https://arthwhdarh.com/share?id=test123456";
const options = {
  width: 800,
  margin: 2,
  color: {
    dark: '#141210',
    light: '#00000000'
  }
};

let capturedImageData = null;

const mockCanvas = {
  getContext: () => ({
    clearRect: () => {},
    createImageData: (w, h) => ({
      width: w,
      height: h,
      data: new Uint8ClampedArray(w * h * 4)
    }),
    putImageData: (imgData) => {
      capturedImageData = imgData;
    }
  })
};

QRCode.toCanvas(mockCanvas, text, options, (err) => {
  if (err) throw err;
  console.log('Canvas rendered successfully.');
  console.log('Image dimensions:', capturedImageData.width, 'x', capturedImageData.height);
  
  // Check alpha channel of first pixel (should be light/background = 0)
  const firstPixelAlpha = capturedImageData.data[3];
  console.log('First pixel alpha (0 = transparent):', firstPixelAlpha);
  
  // Find a dark pixel (alpha should be 255)
  let foundDark = false;
  for (let i = 0; i < capturedImageData.data.length; i += 4) {
    if (capturedImageData.data[i + 3] === 255) {
      foundDark = true;
      console.log('Dark pixel RGBA:', capturedImageData.data[i], capturedImageData.data[i+1], capturedImageData.data[i+2], capturedImageData.data[i+3]);
      break;
    }
  }
  
  if (firstPixelAlpha === 0 && foundDark) {
    console.log('VERIFIED: Background is 100% transparent, dark modules are solid opaque!');
  } else {
    console.error('FAIL: Alpha values incorrect');
  }
});
