import asyncio
from contextlib import contextmanager
import logging
from typing import Optional
import base64
from datetime import datetime
import os

import win32com.client
import pythoncom

from fastapi import FastAPI, HTTPException
from fastapi.responses import JSONResponse
from fastapi import status as fastapi_status
from pydantic import BaseModel

from enums import SignatureEnum

# Логирование.
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
handler = logging.StreamHandler()
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)


class SignRequest(BaseModel):
    data: str
    serial_number: str


class SignResponse(BaseModel):
    signed_data: str
    status: str


class HelperSignature:
    """Помощник для создания подписей на Windows.
    """
    
    @staticmethod
    @contextmanager
    def _com_context():
        # Инициализация COM-библиотеки.
        pythoncom.CoInitialize()
        try:
            yield
        finally:
            pythoncom.CoUninitialize()
    
    @staticmethod
    def _create_com_object(name: str):
        return win32com.client.Dispatch(name)

    def signed_data(self, data: str, serial_number: str) -> str:
        """
        Подписываем данные прикреплённой подписью через COM.

        Args:
            data (str): Готовые данные для подписи (ожидается base64 строка).
            serial_number (str): Серийный номер подписи.

        Returns:
            str: Строка зашифрованных данных прикреплённой подписи (base64).
        """
        with self._com_context():
            logger.info(f"Начинаю поиск сертификата по серийному номеру: {serial_number}")
            # Создаём объект для входа в хранилище подписей.
            oStore = self._create_com_object('CAdESCOM.STORE')
            # Заходим в хранилище подписей.
            oStore.Open(
                SignatureEnum.CAPICOM_CURRENT_USER_STORE.value,
                SignatureEnum.CAPICOM_MY_STORE.value,
                SignatureEnum.CAPICOM_STORE_OPEN_MAXIMUM_ALLOWED.value
            )

            oCert = None

            # Ищем подходящий объект сертификата по серийному номеру.
            for element_oStore in oStore.Certificates:
                cert_serial = element_oStore.SerialNumber.lower()
                if cert_serial == serial_number.lower():
                    print(cert_serial, serial_number.lower())
                    logger.info(f"Найден сертификат: {element_oStore.SubjectName}")
                    oCert = element_oStore
                    break

            if not oCert:
                raise ValueError(f"Подходящий сертификат с серийным номером {serial_number} не найден!")

            # Создаём подписавшего.
            oSigner = self._create_com_object("CAdESCOM.CPSigner")
            oSigner.Certificate = oCert

            # Создаём дополнительные параметры (например, время подписания).
            oSigningTimeAttr = self._create_com_object("CAdESCOM.CPAttribute")
            oSigningTimeAttr.Name = SignatureEnum.TIMER_ATTR_NAME.value
            # Используем datetime.now() вместо Django timezone.now()
            oSigningTimeAttr.Value = datetime.now()

            # Прикрепляем дополнительные параметры подписавшему.
            oSigner.AuthenticatedAttributes2.Add(oSigningTimeAttr)

            # Создаём объект для подписи данных с необходимой кодировкой.
            oSignedData = self._create_com_object("CAdESCOM.CadesSignedData")
            oSignedData.ContentEncoding = SignatureEnum.CONTENT_ENCODING.value

            # Прикрепляем данные к объекту для их подписи.
            oSignedData.Content = data
            # Подписываем.
            logger.info("Начинаю процесс подписания через COM...")
            sSignedData = oSignedData.SignCades(
                oSigner, SignatureEnum.CADES_BES.value,
                False, SignatureEnum.CAPICOM_ENCODE_BASE64.value
            )
            logger.info("Подписание через COM завершено.")

            return sSignedData

    def attached_signed_data(self, data: str, serial_number: str) -> str:
        """
        Подписываем данные прикреплённой подписью.
        Используется для получения динамического токена.
        """
        logger.info("Вызов attached_signed_data")
        # Переводим данные в байты.
        data_bytes = data.encode('ascii')
        # Переводим данные в 64 и обратно декодируем в ascii.
        base64_bytes = base64.b64encode(data_bytes)
        base64_data = base64_bytes.decode('ascii')
        return self.signed_data(base64_data, serial_number)

    def unpinned_signed_data(self, data: str, serial_number: str) -> str:
        """
        Подписываем данные откреплённой подписью.
        Используется для отправки отчётов, проверки кодов и т.д.
        """
        logger.info("Вызов unpinned_signed_data")
        # Строка JSON без пробелов и знаков переносов.
        data = str(data).replace(' ', '\u0020').replace('\n', '').replace('\r', '')
        data_bytes = data.encode()
        base64_bytes = base64.b64encode(data_bytes)
        base64_data = base64_bytes.decode()
        return self.signed_data(base64_data, serial_number)


# FactAPI.
app = FastAPI(title='Signature Service', description='Микросервис для работы с подписями через win32com')

# Инициализация помощника подписей.
helper = HelperSignature()

# Отображение информации.
@app.get("/")
async def root():
    return JSONResponse({
        "service": "Signature Service",
        "version": "1.0",
        "endpoints": {
            "sign_attached": "/api/sign/attached (POST)",
            "sign_unpinned": "/api/sign/unpinned (POST)",
            "docs": "/docs",
            "redoc": "/redoc",
            "openapi_schema": "/openapi.json"
        }
    })

# Эндпоинты.
@app.post('/api/sign/attached/', response_model=SignResponse)
async def sign_attached(request: SignRequest):
    """Подписать данные прикреплённой подписью.

    Args:
        request (SignResponse): Возвращаем подписанные данные и статус.
    """
    
    try:
        logger.info(f'Получен запрос на прикреплённую подпись для серийного номера: {request.serial_number}.')
        signed_data = helper.attached_signed_data(request.data, request.serial_number)
        logger.info('Прикреплённая подпись успешно создана.')
        return SignResponse(signed_data=signed_data, status='succes')
    
    except ValueError as ve:
        logger.error(f'Ошибка валидаци при подписании прикреплённой подписью: {ve}')
        raise HTTPException(
            status_code=fastapi_status.HTTP_404_NOT_FOUND,
            detail=f'Ошибка при создании прикреплённой подписи: {str(ve)}'
        )
    
    except Exception as e:
        logger.error(f'Общая ошибка при подписании прикреплённой подписью: {e}')
        raise HTTPException(
            status_code=fastapi_status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f'Общая ошибка при подписании прикреплённой подписью: {str(e)}'
        )    


@app.post('/api/sign/unpinned/', response_model=SignResponse)
async def sign_unpinned(request: SignRequest):
    """Подписать данные откреплённой подписью.

    Args:
        request (SignRequest): Возвращаем подписанные данные и статус.
    """
    try:
        logger.info(f'Получен запрос на откреплённую подпись для серийника: {request.serial_number}')
        signed_data = helper.unpinned_signed_data(request.data, request.serial_number)
        logger.info('Откреплённая подпись успешно создана.')
        return SignResponse(signed_data=signed_data, status='success')

    except ValueError as ve:
        logger.error(f'Ошибка валидации при подписании откреплённой подписью: {ve}')
        raise HTTPException(
            status_code=fastapi_status.HTTP_404_NOT_FOUND,
            detail=f'Ошибка валидации при подписании откреплённой подписью: {str(ve)}'
        )

    except Exception as e:
        logger.error(f"Общая ошибка при подписании (откреплённая): {e}")
        raise HTTPException(
            status_code=fastapi_status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f'Ошибка при подписании откреплённой подписью: {str(e)}'
        )

# Точка запуска.
if __name__ == '__main__':
    import uvicorn
    HOST = os.getenv('SERVICE_HOST', '0.0.0.0')
    PORT = int(os.getenv('SERVICE_PORT', 8101))
    uvicorn.run(app=app, host=HOST, port=PORT, log_level='debug')
