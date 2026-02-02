import asyncio
from winrt.windows.media.ocr import OcrEngine
from winrt.windows.globalization import Language
from winrt.windows.storage.streams import DataWriter
from winrt.windows.graphics.imaging import SoftwareBitmap, BitmapPixelFormat

def recognize_bytes(bytes, width, height, lang='en'):
    cmd = 'Add-WindowsCapability -Online -Name "Language.OCR~~~en-US~0.0.1.0"'
    assert OcrEngine.is_language_supported(Language(lang)), cmd
    writer = DataWriter()
    writer.write_bytes(bytes)
    sb = SoftwareBitmap.create_copy_from_buffer(writer.detach_buffer(), BitmapPixelFormat.RGBA8, width, height)
    return OcrEngine.try_create_from_language(Language(lang)).recognize_async(sb)

def recognize_pil(img, lang='en'):
    if img.mode != 'RGBA':
        img = img.convert('RGBA')
    return recognize_bytes(img.tobytes(), img.width, img.height, lang)

def recognize_cv2(img, lang='en'):
    import cv2
    img = cv2.cvtColor(img, cv2.COLOR_BGR2RGBA)
    return recognize_bytes(img.tobytes(), img.shape[1], img.shape[0], lang)

def dump_rect(rect):
    return {
        'x': rect.x,
        'y': rect.y,
        'width': rect.width,
        'height': rect.height
    }

def dump_ocrword(word):
    return {
        'bounding_rect': dump_rect(word.bounding_rect),
        'text': word.text
    }

def dump_ocrline(line):
    words = list(map(dump_ocrword, line.words))
    return {
        'text': line.text,
        'words': words
    }

def dump_ocrresult(ocrresult):
    lines = list(map(dump_ocrline, ocrresult.lines))
    return {
        'text': ocrresult.text,
        'text_angle': ocrresult.text_angle,
        'lines': lines
    }

async def to_coroutine(awaitable):
    return await awaitable

def recognize_pil_sync(img, lang='en'):
    return dump_ocrresult(asyncio.run(to_coroutine(recognize_pil(img, lang))))

def recognize_cv2_sync(img, lang='en'):
    return dump_ocrresult(asyncio.run(to_coroutine(recognize_cv2(img, lang))))

def serve():
    import json
    import uvicorn
    from PIL import Image
    from io import BytesIO
    from fastapi import FastAPI, Request, Response
    from fastapi.middleware.cors import CORSMiddleware

    app = FastAPI()
    app.add_middleware(CORSMiddleware, allow_origins=['*'], allow_credentials=True, allow_methods=['*'], allow_headers=['*'])
    @app.post('/')
    async def recognize(request: Request, lang: str = 'en'):
        result = await recognize_pil(Image.open(BytesIO(await request.body())), lang)
        return Response(json.dumps(dump_ocrresult(result), indent=2, ensure_ascii=False), media_type='application/json')
    uvicorn.run(app, host='0.0.0.0')

if __name__ == '__main__':
    serve()
