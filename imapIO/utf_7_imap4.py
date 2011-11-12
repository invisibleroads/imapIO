'Codec for utf-7-imap4 adapted from Twisted'
import codecs


CODEC_NAME = 'utf-7-imap4'


class StreamWriter(codecs.StreamWriter): # pragma: no cover

    def encode(self, s, errors='strict'):
        return encode(s)


class StreamReader(codecs.StreamReader): # pragma: no cover

    def decode(self, s, errors='strict'):
        return decode(s)


def encode(s, errors=None):
    results, encodes = [], []
    def process(units):
        if not units:
            return ''
        encoded = encode_mb64(''.join(units))
        del units[:]
        return '&' + encoded + '-'
    for c in s:
        if ord(c) in range(0x20, 0x26) + range(0x27, 0x7f):
            results.append(process(encodes) + str(c))
        elif c == '&':
            results.append(process(encodes) + '&-')
        else:
            encodes.append(c)
    results.append(process(encodes))
    return ''.join(results), len(s)


def decode(s, errors=None):
    results, decodes = [], []
    def process(units):
        if not units:
            return ''
        decoded = decode_mb64(''.join(units[1:])) if len(units) > 1 else '&'
        del units[:]
        return decoded
    for c in s:
        if c == '&' and not decodes:
            decodes.append('&')
        elif c == '-' and decodes:
            results.append(process(decodes))
        elif decodes:
            decodes.append(c)
        else:
            results.append(c)
    results.append(process(decodes))
    return ''.join(results), len(s)


def encode_mb64(s):
    s_utf7 = s.encode('utf-7')
    return s_utf7[1:-1].replace('/', ',')


def decode_mb64(s):
    s_utf7 = '+' + s.replace(',', '/') + '-'
    return s_utf7.decode('utf-7')


codecPack = encode, decode, StreamReader, StreamWriter
try:
    codecPack = codecs.CodecInfo(*codecPack, name=CODEC_NAME)
except AttributeError: # pragma: no cover
    pass


def utf_7_imap4(codecName):
    if codecName == CODEC_NAME:
        return codecPack


codecs.register(utf_7_imap4)
