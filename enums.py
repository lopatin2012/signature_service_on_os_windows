from enum import Enum

class SignatureEnum(Enum):
    # Параметры подписавшего.
    TIMER_ATTR_NAME = 0
    CONTENT_ENCODING = 1
    
    # Параметры для работы с сертификатом.
    CADES_BES = 1
    CADES_DEFAULT = 0
    CAPICOM_ENCODE_BASE64 = 0
    CAPICOM_CURRENT_USER_STORE = 2
    CAPICOM_MY_STORE = 'My'
    CAPICOM_STORE_OPEN_MAXIMUM_ALLOWED = 2
