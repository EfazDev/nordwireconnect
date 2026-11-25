"""
PyKits v1.9.5 | Made by Efaz from efaz.dev

A usable set of classes with extra functions that can be used within apps.

Import from file:
```python
import PyKits
pip_class = PyKits.pip()
terminal = PyKits.Terminal()
```
"""

# Module Information
__version__ = "1.9.5"
__license__ = "MIT"
__author__ = "EfazDev"
__maintainer__ = "EfazDev"
__email__ = "support@efaz.dev"
__all__ = [
    "requests",
    "Request",
    "Requests",
    "request", 
    "Pip",
    "pip", 
    "Plist",
    "plist", 
    "Terminal", 
    "Translator", 
    "Socket",
    "Lock",
    "ProgressBar",
    "TimerBar",
    "InstantRequestJSONResponse",
    "BuiltinEditor",
    "IterableSetup",
    "FileSelector"
]

# Modules
import typing
import json
import os
import ssl
import sys
import stat
import locale
import builtins
import shutil
import math
import time
import socket
import base64
import zipfile
import functools
import collections
import subprocess
import urllib.parse
import urllib.error
import http.cookiejar
import importlib.util
import importlib.metadata
import importlib
import platform
import ctypes
import uuid
import logging
import zlib
import tempfile
import re
import glob
import threading
import weakref
import mmap
import plistlib
import errno
import select
import concurrent.futures
main_os = platform.system()
cur_path = os.path.dirname(os.path.abspath(__file__))
if main_os != "Windows": 
    import pty
    import fcntl
    import termios
else: 
    import msvcrt
    import winreg
    from ctypes import wintypes
import http.client as http_client
from http.cookiejar import CookieJar as _cookie_jar
import urllib.request as urlreq
from collections import defaultdict
try: import brotli # type: ignore
except ImportError: brotli = None
try: import zstandard as zstd # type: ignore
except ImportError: zstd = None

# PyKits Classes
class request:
    """
    A class that allows you to make HTTP requests using the Python runtime.
    """
    _shared_ssl_context = None
    _zstd_decompressor = None

    class Response:
        __slots__ = (
            "_text", "_json", "redirected_urls", "port", "host",
            "status_code", "headers", "http_version", "path", "url", "method",
            "scheme", "encoding", "redirected", "raw_data", "ssl_handshake_time",
            "loading_time", "ssl_context", "ok", "time", "_ipv6", "_ipv4", "_resp_obj"
        )
        def __init__(self):
            self._text: typing.Optional[str] = None
            self._json: typing.Union[typing.Dict, typing.List, None] = None
            self._ipv4: typing.Union[typing.List[str], None] = None
            self._ipv6: typing.Union[typing.List[str], None] = None
            self._resp_obj = None
            self.redirected_urls: typing.List[str] = []
            self.port: int = 0
            self.host: str = ""
            self.status_code: int = 0
            self.headers: typing.Dict[str, str] = {}
            self.http_version: str = ""
            self.path: str = ""
            self.url: str = ""
            self.method: str = ""
            self.scheme: str = ""
            self.encoding: str = ""
            self.redirected: bool = False
            self.raw_data: bytes = b""
            self.ssl_handshake_time: float = 0
            self.loading_time: float = 0
            self.ssl_context = None
            self.ok: bool = False
            self.time: float = 0.0
        @property
        def ipv4(self):
            if self._ipv4 is None: self._resolve_ips()
            return self._ipv4
        @ipv4.setter
        def ipv4(self, value: typing.List[str]): self._ipv4 = value
        @property
        def ipv6(self):
            if self._ipv6 is None: self._resolve_ips()
            return self._ipv6
        @ipv6.setter
        def ipv6(self, value: typing.List[str]):
            self._ipv6 = value
        @property
        def text(self):
            if self._text is None:
                if self.raw_data: self._text = self.raw_data.decode(errors="replace")
                else: self._text = ""
            return self._text
        @text.setter
        def text(self, value: str):
            self._text = value
        @property
        def json(self):
            if self._json is None and self.text:
                try: self._json = json.loads(self.text)
                except Exception: self._json = None
            return self._json
        @json.setter
        def json(self, value: typing.Union[typing.Dict, typing.List, None]): self._json = value
        def _resolve_ips(self, host: str=None, port: str=None):
            par_url = urllib.parse.urlparse(self.url)
            host = par_url.hostname
            port = 443 if par_url.scheme == "https" else 80
            if not host:
                self._ipv4, self._ipv6 = [], []
                return
            ipv4, ipv6 = [], []
            try:
                addr_info = socket.getaddrinfo(host, port, 0, socket.SOCK_STREAM)
                for family, _, _, _, sockaddr in addr_info:
                    ip = sockaddr[0]
                    if family == socket.AF_INET and ip not in ipv4: ipv4.append(ip)
                    elif family == socket.AF_INET6 and ip not in ipv6: ipv6.append(ip)
            except Exception: pass
            self._ipv4 = ipv4
            self._ipv6 = ipv6
            return ipv4, ipv6
        def iter_content(self, chunk_size: int=8192):
            if self.raw_data:
                for i in range(0, len(self.raw_data), chunk_size): yield self.raw_data[i:i+chunk_size]
            elif hasattr(self, "_resp_obj") and self._resp_obj:
                while True:
                    chunk = self._resp_obj.read(chunk_size)
                    if not chunk:
                        self._resp_obj.close()
                        break
                    yield chunk
    class FileDownload(Response):
        __slots__ = ("returncode", "size", "downloaded")
        def __init__(self):
            super().__init__()
            self.returncode = 0
            self.size = 0
            self.downloaded = 0
            self.path = ""
    class TimedOut(Exception):
        def __init__(self, url: str, time: float): super().__init__(f"Connecting to URL ({url}) took too long to respond in {time}s!")
    class DownloadError(Exception):
        def __init__(self, path: str, status_code: int): super().__init__(f"Unable to download file at {path}! Status code: {status_code}.")
    class UnknownResponse(Exception):
        def __init__(self, url: str, exception: Exception): super().__init__(f"Something went wrong processing the response from URL ({url})! Exception: {str(exception)}")
    class WrongResponse(Exception):
        def __init__(self, url: str): super().__init__(f"Something went wrong processing the response from URL ({url}) because it is not the same type as expected!")
    class ResolveError(Exception):
        def __init__(self, url: str, reason: str): super().__init__(f"Something went wrong trying to resolve the response from URL ({url})! Reason: {reason}")
    class SSLException(Exception):
        def __init__(self, url: str, reason: str): super().__init__(f"Something went wrong to validate SSL trust of ({url})! Reason: {reason}")
    class ConnectionRefusedException(Exception):
        def __init__(self, url: str, reason: str): super().__init__(f"Something went wrong to connect to ({url}) because the connection was refused by the server.")
    class InvalidFileContent(Exception): pass
    class OpenContext:
        __slots__ = ("val",)
        def __init__(self, val): self.val = val
        def __enter__(self): return self.val
        def __exit__(self, exc_type, exc_val, exc_tb): pass
    class DownloadStatus:
        __slots__ = ("speed", "downloaded", "downloaded_bytes", "total_size", "percent")
        def __init__(self, percent: int=0, speed: str="", total_size: str="", downloaded: str="", downloaded_bytes: int=0): 
            self.speed = speed
            self.downloaded = downloaded
            self.percent = percent
            self.downloaded_bytes = downloaded_bytes
            self.total_size = total_size
    class CookieJar:
        __slots__ = ("_cookies",)
        def __init__(self, initial_cookies: typing.Dict[str, str]=None):
            self._cookies = {}
            if initial_cookies:
                for i, v in initial_cookies.items(): self.set(i, v)
        def set(self, key: str, value: str):
            if value == None:
                self._cookies.pop(key, None)
                return
            if not isinstance(key, str) or not isinstance(value, str): raise ValueError("Cookie key and value must be strings.")
            self._cookies[key] = value
        def get(self, key: str): return self._cookies.get(key, None)
        def _generate_http_cookiejar(self, url: str):
            jar = _cookie_jar()
            domain = urllib.parse.urlparse(url).hostname or ""
            is_secure = url.startswith("https://")
            for name, value in self._cookies.items():
                cookie = http.cookiejar.Cookie(
                    version=0,
                    name=name,
                    value=str(value),
                    port=None,
                    port_specified=False,
                    domain=domain,
                    domain_specified=bool(domain),
                    domain_initial_dot=domain.startswith("."),
                    path="/",
                    path_specified=True,
                    secure=is_secure,
                    expires=None,
                    discard=True,
                    comment=None,
                    comment_url=None,
                    rest={},
                    rfc2109=False
                )
                jar.set_cookie(cookie)
            return jar
    __DATA__ = typing.Union[typing.Dict, typing.List, str, bytes]
    __AUTH__ = typing.List[str]
    __HEADERS__ = typing.Dict[str, str]
    __FILES__ = typing.Union[typing.Dict[str, typing.Union[typing.Tuple, typing.List]], typing.Dict[str, str]]
    __COOKIES__ = typing.Union[typing.Dict[str, str], str]
    __slots__ = (
        "cookie_jar", "handle_compression", "automatic_redirect",
        "throw_exceptions", "opener_processors", "_ssl_context", "_site", "_opener_cache", "_conn_pool"
    )
    def __init__(self, handle_compression: bool=True, automatic_redirect: bool=True, opener_processors: typing.List=None, throw_exceptions: bool=True):
        if not opener_processors: opener_processors = []
        self.cookie_jar = _cookie_jar()
        self.handle_compression = handle_compression==True
        self.automatic_redirect = automatic_redirect==True
        self.throw_exceptions = throw_exceptions==True
        self.opener_processors = opener_processors
        if self.__class__._shared_ssl_context is None: self.__class__._shared_ssl_context = self.ensure_python_certs()
        self._ssl_context = self.__class__._shared_ssl_context
        self._opener_cache = {}
        self._conn_pool = {}
    def __bool__(self): return self.get_if_connected()
    def _make_opener(self, jar=None):
        if jar is None: jar = self.cookie_jar
        if not self._ssl_context:
            if self.__class__._shared_ssl_context is None:
                self.__class__._shared_ssl_context = self.ensure_python_certs()
            self._ssl_context = self.__class__._shared_ssl_context
        processor_ids = tuple(id(p) for p in self.opener_processors)
        cache_key = (id(jar), id(self._ssl_context), processor_ids)
        if cache_key in self._opener_cache: return self._opener_cache[cache_key]
        class HTTPStatusProcessor(urlreq.HTTPErrorProcessor):
            def http_response(self, request, response): return response
            https_response = http_response
        processors = []
        processors.append(urlreq.HTTPCookieProcessor(jar))
        processors.append(HTTPStatusProcessor())
        if self._ssl_context: processors.append(urlreq.HTTPSHandler(context=self._ssl_context))
        for i in self.opener_processors:
            if isinstance(i, urlreq.BaseHandler): processors.append(i)
        opener = urlreq.build_opener(*processors)
        self._opener_cache[cache_key] = opener
        return opener
    def _create_response(self, resp, url: str, redirected_urls: typing.List[str]=None, downloading: bool=False, download_path: str="", length: float=0, stream: bool=False):
        if redirected_urls is None: redirected_urls = []
        # Checks
        if not isinstance(resp, http_client.HTTPResponse): raise self.WrongResponse(url)

        # Headers
        res = self.Response() if not downloading else self.FileDownload()
        res.status_code = getattr(resp, "status", None) or resp.getcode()
        res.headers = {i.lower(): v for i, v in resp.headers.items()}
        res.ok = self.get_if_ok(res.status_code)
        res.time = length

        # Fill Data
        res.encoding = res.headers.get("content-encoding")
        if downloading == False:
            data = resp.read()
            if res.encoding: 
                data, res.encoding = self._handle_compression(data, res.encoding)
            res.raw_data = data
        else:
            res.raw_data = b""
            res.returncode = 0 if res.ok else 1
            res.path = download_path

        # Response Info
        res.url = url
        res.path = self.get_url_path(url)
        res.redirected = len(redirected_urls) > 0
        res.redirected_urls = redirected_urls

        # Location
        obj = urllib.parse.urlparse(url)
        res.scheme = obj.scheme
        res.host = obj.netloc
        res.port = obj.port or (443 if obj.scheme == "https" else 80)
        return res
    def _create_blank_response(self, url: str):
        res = self.Response()
        res.url = url
        res.path = self.get_url_path(url)
        res.scheme = self.get_url_scheme(url)
        url_parsed = urllib.parse.urlparse(url)
        res.host = url_parsed.netloc
        res.port = url_parsed.port or (443 if res.scheme == "https" else 80)
        res.status_code = -1
        res.ok = False
        return res
    def _add_auth_to_headers(self, headers: __HEADERS__, auth: __AUTH__):
        if not auth: return headers
        if isinstance(auth, (tuple, list)) and len(auth) == 2:
            user, pw = auth
            token = base64.b64encode(f"{user}:{pw}".encode()).decode()
            headers["Authorization"] = f"Basic {token}"
        elif isinstance(auth, str):
            if not auth.lower().startswith(("basic ", "bearer ")): auth = f"Bearer {auth}"
            headers["Authorization"] = auth
        return headers
    def _make_request(self, url: str, method: str, data: __DATA__, headers: __HEADERS__=None, cookies: __COOKIES__=None, auth: __AUTH__=None, files: __FILES__=None, timeout: float=30.0, follow_redirects: bool=None, loop_429: bool=False, loop_count: int=-1, loop_timeout: int=1):
        # Prepare to make the request
        req_headers = dict(headers) if headers else {}
        if self.handle_compression: req_headers.setdefault("Accept-Encoding", "gzip, deflate")
        if auth is None: auth = []
        if files is None: files = {}
        if follow_redirects == None: follow_redirects = self.automatic_redirect
        url, opener, method, data, req_headers, cookies, auth, files = self._handle_data(url, method=method, data=data, headers=req_headers, cookies=cookies, auth=auth, files=files)
        redirected_urls = []

        # Let's make the request!
        try:
            start = time.perf_counter()
            req = urlreq.Request(url, data=data, method=method, headers=req_headers)
            with opener.open(req, timeout=timeout) as resp:
                ssl_handshake = (time.perf_counter()-start)*1000
                start = time.perf_counter()
                if follow_redirects:
                    status_code = getattr(resp, "status", None) or resp.getcode()
                    redirects_left = 10
                    while self.get_if_redirect(status_code) and redirects_left > 0:
                        redirects_left -= 1
                        redirected_urls.append(url)
                        next_url = urllib.parse.urljoin(url, resp.headers.get("location"))
                        resp.close()
                        url = next_url
                        req = urlreq.Request(url, data=data, method=method, headers=req_headers)
                        resp = opener.open(req, timeout=timeout)
                        status_code = getattr(resp, "status", None) or resp.getcode()
                response_obj = self._create_response(resp, url, redirected_urls=redirected_urls, downloading=False)
                response_obj.ssl_handshake_time = ssl_handshake
                if self.get_if_cooldown(response_obj.status_code) and loop_429 == True and ((1 if loop_count == -1 else loop_count) >= 1):
                    time.sleep(loop_timeout)
                    response_obj = self._make_request(url, method, data=data, headers=headers, cookies=cookies, auth=auth, files=files, timeout=timeout, follow_redirects=True, loop_429=loop_429, loop_count=(loop_count-1 if loop_count != -1 else loop_count), loop_timeout=loop_timeout)
                response_obj.loading_time = (time.perf_counter()-start)*1000
                return response_obj
        except urllib.error.URLError as e:
            if self.throw_exceptions == False: return self._create_blank_response(url)
            if isinstance(e.reason, ssl.SSLCertVerificationError): raise self.SSLException(url, str(e.reason))
            elif isinstance(e.reason, socket.timeout): raise self.TimedOut(url, timeout)
            elif isinstance(e.reason, socket.gaierror): raise self.ResolveError(url, str(e.reason))
            elif isinstance(e.reason, ConnectionRefusedError) or isinstance(e.reason, ConnectionResetError): raise self.ConnectionRefusedException(url, str(e.reason))
            raise self.ResolveError(url, str(e.reason))
        except ssl.SSLError as e:
            if self.throw_exceptions == False: return self._create_blank_response(url)
            raise self.SSLException(url, str(e))
        except getattr(http_client, 'RemoteDisconnected', Exception) as e:
            if self.throw_exceptions == False: return self._create_blank_response(url)
            raise self.ConnectionRefusedException(url, str(e))
        except (ConnectionResetError, BrokenPipeError) as e:
            if self.throw_exceptions == False: return self._create_blank_response(url)
            raise self.ConnectionRefusedException(url, str(e))
        except (socket.timeout, TimeoutError) as e:
            if self.throw_exceptions == False: return self._create_blank_response(url)
            raise self.TimedOut(url, timeout)
        except OSError as e:
            if self.throw_exceptions == False: return self._create_blank_response(url)
            raise self.ResolveError(url, str(e))
        except Exception as e: 
            if self.throw_exceptions == False: return self._create_blank_response(url)
            raise self.UnknownResponse(url, e)
    def _handle_compression(self, data: bytes, encoding: str):
        # Reject Compression Reading if Disabled
        if self.handle_compression == False: return data, encoding
        
        # Format response
        encoding = (encoding or "").lower().strip()
        if not data: return b"", encoding

        # Auto determine encoding
        if not encoding:
            if data.startswith(b"\x1f\x8b"): encoding = "gzip"
            elif data.startswith(b"\x78\x9c") or data.startswith(b"\x78\xda"): encoding = "deflate"
            elif data.startswith(b"\x28\xb5\x2f\xfd"): encoding = "zstd"
            elif brotli and data[:2] in (b"\x8b\x0b", b"\xce\xb2"): encoding = "br"

        # Decompress data
        if encoding == "gzip": return zlib.decompress(data, zlib.MAX_WBITS | 32), encoding
        elif encoding == "deflate":
            try: return zlib.decompress(data, -zlib.MAX_WBITS), encoding
            except zlib.error: return zlib.decompress(data), encoding
        elif encoding == "br" and brotli: return brotli.decompress(data), encoding
        elif encoding == "zstd" and zstd: 
            if not self.__class__._zstd_decompressor: self.__class__._zstd_decompressor = zstd.ZstdDecompressor()
            return self.__class__._zstd_decompressor.decompress(data), encoding
        return data, encoding
    def _handle_data(self, url: str, method: str="GET", data: __DATA__=None, headers: __HEADERS__=None, cookies: __COOKIES__=None, auth: __AUTH__=None, files: __FILES__=None):
        try:
            if headers is None: headers = {}
            if auth is None: auth = []
            if files is None: files = {}
            if not "://" in url: url = "http://" + url
            if isinstance(cookies, self.CookieJar): cookie_jar = cookies._generate_http_cookiejar(url)
            elif isinstance(cookies, dict) and len(cookies) > 0: cookie_jar = self.CookieJar(cookies)._generate_http_cookiejar(url)
            else: cookie_jar = self.cookie_jar
            headers.setdefault("user-agent", f"PyKits/{__version__}")
            headers = self._add_auth_to_headers(headers, auth)
            opener = self._make_opener(jar=cookie_jar)
            method = method.upper()
            if method in ("POST", "PUT", "PATCH") and (isinstance(data, self.__DATA__) or len(files) > 0):
                if len(files) > 0:
                    boundary = f"----WebKitFormBoundary{uuid.uuid4().hex}"
                    body = bytearray() 
                    if isinstance(data, dict):
                        for k, v in data.items(): body.extend(f"--{boundary}\r\nContent-Disposition: form-data; name=\"{k}\"\r\n\r\n{v}\r\n".encode())
                    for field, info in files.items():
                        filename, content, mimetype = None, None, None
                        if isinstance(info, (tuple, list)):
                            filename = info[0]
                            content = info[1]
                            mimetype = info[2] if len(info) > 2 else "application/octet-stream"
                        elif isinstance(info, str):
                            with open(info, "rb") as f: content = f.read()
                            filename = os.path.basename(info)
                            mimetype = "application/octet-stream"
                        else: raise ValueError("File value must be a tuple of (filename, content, mimetype)")
                        body.extend(
                            (f"--{boundary}\r\n"
                            f"Content-Disposition: form-data; name=\"{field}\"; filename=\"{filename}\"\r\n"
                            f"Content-Type: {mimetype}\r\n\r\n").encode()
                        )
                        if isinstance(content, str): content = content.encode()
                        elif not isinstance(content, (bytes, bytearray)): raise self.InvalidFileContent("File content must be bytes or bytearray")
                        body.extend(content)
                        body.extend(b"\r\n")
                    body.extend(f"--{boundary}--\r\n".encode())
                    data = bytes(body) 
                    headers["Content-Type"] = f"multipart/form-data; boundary={boundary}"
                    headers["Content-Length"] = str(len(data))
                elif isinstance(data, str): data = data.encode()
                elif isinstance(data, bytes): pass
                else: data = json.dumps(data).encode(); headers.setdefault("Content-Type", "application/json")
            else: data = None
            return url, opener, method, data, headers, cookies, auth, files
        except Exception as e: raise self.UnknownResponse(url, e)
    def _pkg_check(self, package: str):
        try: 
            importlib.metadata.version(package)
            return True
        except Exception:
            try: return importlib.util.find_spec(package) is not None
            except Exception: return False
    def get(self, url: str, headers: __HEADERS__=None, cookies: __COOKIES__=None, auth: __AUTH__=None, timeout: float=30.0, follow_redirects: bool=None, loop_429: bool=False, loop_count: int=-1, loop_timeout: int=1) -> Response:
        return self._make_request(url=url, method="GET", data=None, headers=headers, cookies=cookies, auth=auth, timeout=timeout, follow_redirects=follow_redirects, loop_429=loop_429, loop_count=loop_count, loop_timeout=loop_timeout)
    def post(self, url: str, data: __DATA__, headers: __HEADERS__=None, cookies: __COOKIES__=None, auth: __AUTH__=None, files: __FILES__=None, timeout: float=30.0, follow_redirects: bool=None, loop_429: bool=False, loop_count: int=-1, loop_timeout: int=1) -> Response:
        return self._make_request(url=url, method="POST", data=data, headers=headers, cookies=cookies, auth=auth, files=files, timeout=timeout, follow_redirects=follow_redirects, loop_429=loop_429, loop_count=loop_count, loop_timeout=loop_timeout)
    def patch(self, url: str, data: __DATA__, headers: __HEADERS__=None, cookies: __COOKIES__=None, auth: __AUTH__=None, files: __FILES__=None, timeout: float=30.0, follow_redirects: bool=None, loop_429: bool=False, loop_count: int=-1, loop_timeout: int=1) -> Response:
        return self._make_request(url=url, method="PATCH", data=data, headers=headers, cookies=cookies, auth=auth, files=files, timeout=timeout, follow_redirects=follow_redirects, loop_429=loop_429, loop_count=loop_count, loop_timeout=loop_timeout)
    def put(self, url: str, data: __DATA__, headers: __HEADERS__=None, cookies: __COOKIES__=None, auth: __AUTH__=None, files: __FILES__=None, timeout: float=30.0, follow_redirects: bool=None, loop_429: bool=False, loop_count: int=-1, loop_timeout: int=1) -> Response:
        return self._make_request(url=url, method="PUT", data=data, headers=headers, cookies=cookies, auth=auth, files=files, timeout=timeout, follow_redirects=follow_redirects, loop_429=loop_429, loop_count=loop_count, loop_timeout=loop_timeout)
    def delete(self, url: str, headers: __HEADERS__=None, cookies: __COOKIES__=None, auth: __AUTH__=None, timeout: float=30.0, follow_redirects: bool=None, loop_429: bool=False, loop_count: int=-1, loop_timeout: int=1) -> Response:
        return self._make_request(url=url, method="DELETE", data=None, headers=headers, cookies=cookies, auth=auth, timeout=timeout, follow_redirects=follow_redirects, loop_429=loop_429, loop_count=loop_count, loop_timeout=loop_timeout)
    def head(self, url: str, headers: __HEADERS__=None, cookies: __COOKIES__=None, auth: __AUTH__=None, timeout: float=30.0, follow_redirects: bool=None, loop_429: bool=False, loop_count: int=-1, loop_timeout: int=1) -> Response:
        return self._make_request(url=url, method="HEAD", data=None, headers=headers, cookies=cookies, auth=auth, timeout=timeout, follow_redirects=follow_redirects, loop_429=loop_429, loop_count=loop_count, loop_timeout=loop_timeout)
    def custom(self, url: str, method: str, data: __DATA__=None, headers: __HEADERS__=None, cookies: __COOKIES__=None, auth: __AUTH__=None, timeout: float=30.0, follow_redirects: bool=None, loop_429: bool=False, loop_count: int=-1, loop_timeout: int=1) -> Response:
        return self._make_request(url=url, method=method, data=data, headers=headers, cookies=cookies, auth=auth, timeout=timeout, follow_redirects=follow_redirects, loop_429=loop_429, loop_count=loop_count, loop_timeout=loop_timeout)
    def open(self, *k, **s) -> OpenContext:
        mai = self._make_request(*k, **s)
        return self.OpenContext(mai)
    def download(self, url: str, output: str, method: str="GET", data: __DATA__=None, headers: __HEADERS__=None, cookies: __COOKIES__=None, auth: __AUTH__=None, files: __FILES__=None, check: bool=False, follow_redirects: bool=None, delete_existing: bool=True, submit_status=None, chunk_size: int=65535, threads: int=1, timeout: float=30.0) -> FileDownload:
        try:
            # Assurance
            headers = dict(headers) if headers else {}
            if auth is None: auth = []
            if files is None: files = {}
            if not self._ssl_context:
                if self.__class__._shared_ssl_context is None: self.__class__._shared_ssl_context = self.ensure_python_certs()
                self._ssl_context = self.__class__._shared_ssl_context
            if os.path.exists(output) and delete_existing == False: raise FileExistsError(f"This file already exists in {output}!")
            elif os.path.exists(output) and os.path.isdir(output): shutil.rmtree(output, ignore_errors=True)
            elif os.path.exists(output) and os.path.isfile(output): os.remove(output)

            # Prepare Data
            redirected_urls = []
            url, opener, method, data, headers, cookies, auth, files = self._handle_data(url, method=method, data=data, headers=headers, cookies=cookies, auth=auth, files=files)
            req = urlreq.Request(url, data=data, method=method, headers=headers)
            start_download = time.perf_counter()

            # Let's download the request!
            if follow_redirects == None: follow_redirects = self.automatic_redirect
            if not chunk_size: chunk_size = 65535
            with opener.open(req, timeout=timeout) as resp:
                ssl_handshake = (time.perf_counter()-start_download)*1000
                start_download = time.perf_counter()
                if follow_redirects:
                    status_code = getattr(resp, "status", None) or resp.getcode()
                    redirects_left = 10
                    while self.get_if_redirect(status_code) and redirects_left > 0:
                        redirects_left -= 1
                        redirected_urls.append(url)
                        next_url = urllib.parse.urljoin(url, resp.headers.get("location"))
                        resp.close()
                        url = next_url
                        req = urlreq.Request(url, data=data, method=method, headers=headers)
                        resp = opener.open(req, timeout=timeout)
                        status_code = getattr(resp, "status", None) or resp.getcode()
                download_info = self._create_response(resp, url, downloading=True, download_path=output)
                if check == True and not download_info.ok: raise self.DownloadError(url, download_info.status_code)
                total = download_info.headers.get("content-length")
                encoding = download_info.headers.get("content-encoding")
                accept_ranges = download_info.headers.get("accept-ranges", "").lower() == "bytes"
                if total is not None: total = int(total)
                chunk_size = (min(chunk_size, total) if total else chunk_size)
                if threads > 1 and accept_ranges and total and not encoding:
                    resp.close()
                    part_size = total // threads
                    ranges = []
                    for i in range(threads):
                        start_byte = i * part_size
                        end_byte = total - 1 if i == threads - 1 else start_byte + part_size - 1
                        ranges.append((start_byte, end_byte))
                    with open(output, "wb") as f:
                        f.truncate(total)
                    progress_lock = threading.Lock()
                    shared_downloaded = 0
                    last_status_time = time.perf_counter()
                    last_downloaded_marker = 0
                    def download_section(part_num, start_b, end_b):
                        nonlocal shared_downloaded, last_status_time, last_downloaded_marker
                        part_headers = headers.copy()
                        part_headers["Range"] = f"bytes={start_b}-{end_b}"
                        part_req = urlreq.Request(url, data=data, method=method, headers=part_headers)
                        local_chunk_count = 0
                        with opener.open(part_req) as part_resp:
                            with open(output, "r+b") as f:
                                f.seek(start_b)
                                while True:
                                    chunk = part_resp.read(chunk_size)
                                    if not chunk: break
                                    f.write(chunk)
                                    local_chunk_count += len(chunk)
                                    now = time.perf_counter()
                                    if local_chunk_count > 1048576 or (now - last_status_time) >= 0.1:
                                        with progress_lock:
                                            shared_downloaded += local_chunk_count
                                            local_chunk_count = 0
                                            if submit_status and (now - last_status_time) >= 0.1:
                                                duration = now - last_status_time
                                                speed = ((shared_downloaded - last_downloaded_marker) / duration) if duration > 0 else 0.0
                                                last_status_time, last_downloaded_marker = now, shared_downloaded
                                                percent = (shared_downloaded * 100 / total)
                                                progress = self.DownloadStatus(speed=speed, downloaded=self.format_bytes_to_size(shared_downloaded), downloaded_bytes=shared_downloaded, percent=percent, total_size=total)
                                                submit_status.submit(progress)
                                if local_chunk_count > 0:
                                    with progress_lock: shared_downloaded += local_chunk_count
                    with concurrent.futures.ThreadPoolExecutor(max_workers=threads) as executor:
                        futures = [executor.submit(download_section, i, r[0], r[1]) for i, r in enumerate(ranges)]
                        concurrent.futures.wait(futures)
                        for future in futures: future.result()
                    download_info.downloaded = shared_downloaded
                    download_info.size = total
                else:
                    with open(output, "ab") as f:
                        downloaded = 0
                        last_status = 0.0
                        decompressor = None
                        if encoding == "gzip": decompressor = zlib.decompressobj(zlib.MAX_WBITS | 32)
                        elif encoding == "deflate":
                            try: decompressor = zlib.decompressobj(zlib.MAX_WBITS)
                            except Exception: decompressor = zlib.decompressobj(-zlib.MAX_WBITS)
                        elif encoding == "br" and brotli: decompressor = brotli.Decompressor()
                        elif encoding == "zstd" and zstd: decompressor = zstd.ZstdDecompressor().decompressobj()
                        while True:
                            # Read and write chunk into file
                            start = time.perf_counter()
                            chunk = resp.read(chunk_size)
                            downloaded_chunk_size = len(chunk)
                            if not chunk:
                                if decompressor and hasattr(decompressor, "flush"):
                                    flushed = decompressor.flush()
                                    if flushed: f.write(flushed)
                                break
                            if decompressor:
                                try: chunk = decompressor.decompress(chunk)
                                except zlib.error:
                                    decompressor = zlib.decompressobj(-zlib.MAX_WBITS)
                                    chunk = decompressor.decompress(chunk)
                            f.write(chunk)
                            end = time.perf_counter()
                            duration = end-start
                            speed = (downloaded_chunk_size / duration) if duration > 0 else 0.0

                            # Update status
                            downloaded += downloaded_chunk_size
                            now = time.perf_counter()
                            if submit_status and (now - last_status) >= 0.1:
                                last_status = now
                                percent = (downloaded * 100 / total) if total else (99.9 if downloaded_chunk_size > 0 else 100.0)
                                progress = self.DownloadStatus(speed=speed, downloaded=self.format_bytes_to_size(downloaded), downloaded_bytes=downloaded, percent=percent, total_size=total)
                                submit_status.submit(progress)
                        download_info.downloaded = downloaded
                        if not total: total = os.path.getsize(output)
                        download_info.size = total
                download_info.ssl_handshake_time = ssl_handshake
                download_info.loading_time = (time.perf_counter()-start_download)*1000
                return download_info
        except urllib.error.URLError as e:
            if isinstance(e.reason, ssl.SSLCertVerificationError): raise self.SSLException(url, str(e.reason))
            elif isinstance(e.reason, socket.gaierror): raise self.ResolveError(url, str(e.reason))
            elif isinstance(e.reason, ConnectionRefusedError) or isinstance(e.reason, ConnectionResetError): raise self.ConnectionRefusedException(url, str(e.reason))
            raise self.ResolveError(url, str(e.reason))
        except Exception as e: raise self.DownloadError(url, str(e))
    def ensure_python_certs(self, certifi_only: bool=False):
        ssl_ctx = None
        is_compiled = bool(getattr(sys, "frozen", False) or "__compiled__" in globals())
        alleged_path = os.path.dirname(sys.executable)
        virt = os.path.exists(os.path.join(alleged_path, "..", "pyvenv.cfg")) or (os.path.exists(os.path.join(alleged_path, "python.exe")) and os.path.exists(os.path.join(alleged_path, "pip.exe")))
        if certifi_only == False and sys.version_info >= (3, 10, 0): 
            try:
                try:
                    import truststore # type: ignore
                    return truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
                except ImportError:
                    if is_compiled: return self.ensure_python_certs(certifi_only=True)
                if self._pkg_check("truststore") == False:
                    import site
                    self._site = site
                    s = subprocess.run([sys.executable, "-m", "pip", "install"] + (["--user"] if (not virt and self._site.ENABLE_USER_SITE) else []) + ["--upgrade", "truststore"], stdout=subprocess.DEVNULL)
                    if s.returncode == 0:
                        if not virt:
                            site_packages_paths = self._site.getsitepackages() + [self._site.getusersitepackages()]
                            for path in site_packages_paths:
                                if path not in sys.path and os.path.exists(path): sys.path.append(path)
                        importlib.invalidate_caches()
                    else: return self.ensure_python_certs(certifi_only=True)
                import truststore # type: ignore
                ssl_ctx = truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
            except Exception: return self.ensure_python_certs(certifi_only=True)
        elif not is_compiled and self._pkg_check("certifi") == False:
            STAT_0o775 = ( stat.S_IRUSR | stat.S_IWUSR | stat.S_IXUSR | stat.S_IRGRP | stat.S_IWGRP | stat.S_IXGRP | stat.S_IROTH |  stat.S_IXOTH )
            openssl_dir, openssl_cafile = os.path.split(ssl.get_default_verify_paths().openssl_cafile)
            subprocess.check_call([sys.executable, "-E", "-m", "pip", "install"] + (["--user"] if not virt else []) + ["--upgrade", "certifi"])
            import certifi
            os.chdir(openssl_dir)
            relpath_to_certifi_cafile = os.path.relpath(certifi.where())
            try: os.remove(openssl_cafile)
            except FileNotFoundError: pass
            os.symlink(relpath_to_certifi_cafile, openssl_cafile)
            os.chmod(openssl_cafile, STAT_0o775)
            ssl_ctx = ssl.create_default_context()
        else:
            try:
                import certifi # type: ignore
                ssl_ctx = ssl.create_default_context(cafile=certifi.where())
            except Exception: ssl_ctx = ssl.create_default_context()
        return ssl_ctx
    def get_if_ok(self, code: int): return int(code) < 300 and int(code) >= 200
    def get_if_redirect(self, code: int): return int(code) < 400 and int(code) >= 300
    def get_if_ip(self, ip: int):
        try: socket.inet_pton(socket.AF_INET, ip); return True
        except OSError: 
            try: socket.inet_pton(socket.AF_INET6, ip); return True
            except Exception: return False
        except Exception: return False
    def get_if_cooldown(self, code: int): return int(code) == 429
    def generate_location_url(self, location_header: str="", host: str="https://google.com"):
        if not location_header: return ""
        if location_header.startswith("/"): return host + location_header
        return location_header
    def get_if_connected(self, server: str="8.8.8.8", port: int=443, timeout: int=3):
        try: socket.create_connection((server, port), timeout=timeout).close(); return True # Connect to server failed = Disconnected
        except Exception: return False
    def get_url_scheme(self, url: str): 
        obj = urllib.parse.urlparse(url)
        return obj.scheme
    def get_url_path(self, url: str):
        obj = urllib.parse.urlparse(url)
        if obj.query == "": return obj.path
        else: return obj.path + "?" + obj.query
    def format_params(self, data: typing.Dict[str, str]=None):
        if not data: return ""
        return "?" + urllib.parse.urlencode(data)
    def format_size_to_bytes(self, size_str: str):
        size_str = size_str.upper()
        try:
            if size_str.endswith("K") or size_str.endswith("k"): return int(float(size_str[:-1]) * 1024)
            if size_str.endswith("M"): return int(float(size_str[:-1]) * 1024**2)
            if size_str.endswith("G"): return int(float(size_str[:-1]) * 1024**3)
            if size_str.endswith("T"): return int(float(size_str[:-1]) * 1024**4)
            return int(size_str)
        except Exception: return 0
    def format_bytes_to_size(self, size_bytes: int):
        thresholds = [
            (1024**4, "T"),
            (1024**3, "G"),
            (1024**2, "M"),
            (1024, "k"),
        ]
        for factor, suffix in thresholds:
            if size_bytes >= factor:
                size = size_bytes / factor
                return f"{size:.1f}{suffix}"
        return str(size_bytes)
    def process_bytes_to_str(self, bytes: bytes): return bytes.decode("utf-8", errors="ignore")
class pip:
    """
    A class that allows you to configure pip and Python installations.
    """
    _class_cache_initialized = False
    _uv_exe_path = None
    _brew_exe_path = None
    _latest_python_cache = {}
    _python_path_cache = {}
    __slots__ = (
        "executable", "debug", "ignore_same", "_requests_obj", "_daemon_threads", 
        "iter_index", "iter_data", "_win32con", "_win32api", "win32con", 
        "_win32gui", "_win32process", "_site", "_CGWindowListCopyWindowInfo", 
        "_kCGWindowListOptionOnScreenOnly", "uv_exe", "brew_exe", 
        "_pip_ensured", "_is_same_env", "_arch_cache", "_current_version_cache"
    )
    
    # Pip / PyPi Functionalities
    def __init__(self, command: typing.List=None, executable: str=None, debug: bool=False, find: bool=False, arch: str=None):
        if not self.__class__._class_cache_initialized:
            self.__class__._uv_exe_path = shutil.which("uv")
            self.__class__._brew_exe_path = shutil.which("brew")
            self.__class__._class_cache_initialized = True
        self._daemon_threads = weakref.WeakSet()
        self.ignore_same = False
        self.uv_exe = self.__class__._uv_exe_path
        self.brew_exe = self.__class__._brew_exe_path
        self._requests_obj = None
        self._pip_ensured = False
        self._is_same_env = None
        self._arch_cache = None
        self._current_version_cache = None
        if command is None: command = []
        if isinstance(executable, str):
            if os.path.isfile(executable): self.executable = executable
            else: self.executable = self.findPython(arch=arch, path=True) if find == True else sys.executable
        elif isinstance(arch, str): self.executable = self.findPython(arch=arch, path=True)
        else: self.executable = self.findPython(arch=arch, path=True) if find == True else sys.executable
        self.debug = debug==True
        if isinstance(command, list) and len(command) > 0: self.ensure(); subprocess.check_call([self.executable, "-m", "pip"] + command)
    @property
    def requests(self):
        if self._requests_obj is None: self._requests_obj = request()
        return self._requests_obj
    def __str__(self): return self.executable
    def __bool__(self): return self.pythonInstalled()
    def __iter__(self): return self
    def __next__(self):
        if not hasattr(self, "iter_index") or not self.iter_index:
            self.iter_index = 0
            self.iter_data = self.findPythons()
        if self.iter_index < len(self.iter_data):
            item = self.iter_data[self.iter_index]
            self.iter_index += 1
            return item
        else:
            self.iter_data = None
            self.iter_index = 0
            raise StopIteration
    def _pkg_check(self, package: str):
        try: 
            importlib.metadata.version(package)
            return True
        except Exception:
            try: return importlib.util.find_spec(package) is not None
            except Exception: return False
    def _to_int(self, val): return int("".join(filter(str.isdigit, val)) or 0)
    @classmethod
    def clear_class_caches(cls):
        cls._class_cache_initialized = False
        cls._uv_exe_path = None
        cls._brew_exe_path = None
        cls._latest_python_cache.clear()
        cls._python_path_cache.clear()
    def refresh_cache(self):
        self._pip_ensured = False
        self._is_same_env = None
        self._arch_cache = None
        self._current_version_cache = None
        self.clear_class_caches()
        importlib.invalidate_caches()
        return self
    def install(self, packages: typing.List[str], upgrade: bool=False, user: bool=True):
        self.ensure()
        res = {}
        generated_list = [i for i in packages if isinstance(i, str)]
        if self.getIfVirtualEnvironment(): user = False
        if len(generated_list) > 0:
            try:
                if self.uv_exe: cmd = [self.uv_exe, "pip", "install", "--python", self.executable] + (["--upgrade"] if upgrade else []) + generated_list
                else: cmd = [self.executable, "-m", "pip", "install"] + (["--upgrade"] if upgrade else []) + (["--user"] if user else []) + generated_list
                a = subprocess.call(cmd, stdout=(not self.debug) and subprocess.DEVNULL or None, stderr=(not self.debug) and subprocess.DEVNULL or None)
                if a == 0: 
                    importlib.invalidate_caches()
                    return {"success": True, "message": "Successfully installed modules!"}
                else: return {"success": False, "message": f"Command has failed! Code: {a}"}
            except Exception as e: return {"success": False, "message": str(e)}
        return res
    def uninstall(self, packages: typing.List[str]):
        self.ensure()
        res = {}
        generated_list = [i for i in packages if isinstance(i, str)]
        if len(generated_list) > 0:
            try:
                if self.uv_exe: cmd = [self.uv_exe, "pip", "uninstall", "--python", self.executable] + generated_list
                else: cmd = [self.executable, "-m", "pip", "uninstall", "-y"] + generated_list
                a = subprocess.call(cmd, stdout=subprocess.DEVNULL if self.debug == False else None, stderr=subprocess.DEVNULL if self.debug == False else None)
                importlib.invalidate_caches()
                for i in generated_list: res[i] = {"success": True}
            except Exception as e:
                for i in generated_list: res[i] = {"success": False}
        return res
    def installed(self, packages: typing.List[str]=[], boolonly: bool=False):
        if self.isSameRunningPythonExecutable():
            if len(packages) == 0: return [dist.metadata['Name'].lower() for dist in importlib.metadata.distributions()]  
            if len(packages) == 1: return self._pkg_check(packages[0].lower())
            else:
                installed_checked = {}
                all_installed = True
                for i in packages:
                    check = self._pkg_check(i.lower())
                    installed_checked[i] = check
                    if not check: all_installed = False
                installed_checked["all"] = all_installed
                if boolonly == True: return installed_checked["all"]
                return installed_checked
        else:
            self.ensure()
            if len(packages) == 1:
                code = f"import importlib.util, sys; sys.exit(0 if importlib.util.find_spec('{packages[0]}') else 1)"
                return subprocess.run([self.executable, "-c", code], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL).returncode == 0
            sub = subprocess.run([self.executable, "-m", "pip", "list"], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            line_splits = sub.stdout.decode().strip().splitlines()[2:]
            installed_packages = [package.split()[0].lower() for package in line_splits if package.strip()]
            installed_checked = {}
            all_installed = True
            if len(packages) == 0: return installed_packages
            else:
                for i in packages:
                    if i.lower() in installed_packages: installed_checked[i] = True
                    else:
                        installed_checked[i] = False
                        all_installed = False
                installed_checked["all"] = all_installed
                if boolonly == True: return installed_checked["all"]
                return installed_checked
    def download(self, packages: typing.List[str], repository_mode: bool=False):
        generated_list = [i for i in packages if isinstance(i, str)]
        if len(generated_list) > 0:
            try:
                if repository_mode == True:
                    url_paths, url_paths_2 = [], []
                    for i in generated_list: 
                        if i.startswith("https://github.com") or i.startswith("https://www.github.com"):
                            path_parts = urllib.parse.urlparse(i).path.strip('/').split('/')
                            url_paths.append(path_parts[-1])
                            url_paths_2.append(path_parts[-2])
                    down_path = os.path.join(cur_path, '-'.join(url_paths) + "_download")
                    if os.path.isdir(down_path): shutil.rmtree(down_path, ignore_errors=True)
                    os.makedirs(down_path)
                    os.chmod(down_path, 0o755)
                    downed_paths = []
                    for url_path_1, url_path_2 in zip(url_paths, url_paths_2):
                        s = self.requests.download(f"https://github.com/{url_path_2}/{url_path_1}/archive/refs/heads/main.zip", os.path.join(down_path, f"{url_path_1}.zip"))
                        if s.ok: downed_paths.append(os.path.join(down_path, f"{url_path_1}.zip"))
                    return {"success": True, "path": down_path, "package_files": downed_paths}
                else:
                    down_path = os.path.join(cur_path, '-'.join(generated_list) + "_download")
                    if os.path.isdir(down_path): shutil.rmtree(down_path, ignore_errors=True)
                    os.makedirs(down_path)
                    os.chmod(down_path, 0o755)
                    self.ensure()
                    subprocess.check_call([self.executable, "-m", "pip", "download", "--no-binary", ":all:"] + generated_list, stdout=self.debug == False and subprocess.DEVNULL, stderr=self.debug == False and subprocess.DEVNULL, cwd=down_path)
                    a = [os.path.join(down_path, e) for e in os.listdir(down_path)]
                    return {"success": True, "path": down_path, "package_files": a}
            except Exception as e:
                print(e)
                return {"success": False}
        return {"success": False}
    def update(self):
        self.ensure()
        try:
            if self.uv_exe: cmd = [self.uv_exe, "pip", "install", "--python", self.executable, "--upgrade", "pip"]
            else: cmd = [self.executable, "-m", "pip", "install", "--upgrade", "pip"]
            a = subprocess.call(cmd, stdout=subprocess.DEVNULL if (not self.debug) else None, stderr=subprocess.DEVNULL if (not self.debug) else None)
            if a == 0: return {"success": True, "message": "Successfully installed latest version of pip!"}
            else: return {"success": False, "message": f"Command has failed!"}
        except Exception as e: return {"success": False, "message": str(e)}
    def ensure(self):
        if not self.executable: return False
        if self._pip_ensured: return True
        check_for_pip_pro = subprocess.run([self.executable, "-m", "pip"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        if check_for_pip_pro.returncode == 0: 
            self._pip_ensured = True
            return True
        else:
            if self.getIfConnectedToInternet() == True:
                self.printDebugMessage(f"Downloading pip from pypi..")
                with tempfile.NamedTemporaryFile(suffix=".py", delete=False) as temp_file: pypi_download_path = temp_file.name
                if self.pythonSupported(3,9,0): download_res = self.requests.download("https://bootstrap.pypa.io/get-pip.py", pypi_download_path)      
                else: current_python_version = self.getCurrentPythonVersion(); download_res = self.requests.download(f"https://bootstrap.pypa.io/pip/{current_python_version.split('.')[0]}.{current_python_version.split('.')[1]}/get-pip.py", pypi_download_path)
                if download_res.ok:
                    self.printDebugMessage(f"Successfully downloaded pip! Installing to Python..")
                    install_to_py = subprocess.run([self.executable, pypi_download_path], stdout=self.debug == False and subprocess.DEVNULL, stderr=self.debug == False and subprocess.DEVNULL)
                    if install_to_py.returncode == 0:
                        self.printDebugMessage(f"Successfully installed pip to Python executable!")
                        self._pip_ensured = True
                        return True
                    else: return False
                else: return False
            else:
                self.printDebugMessage(f"Unable to download pip due to no internet access.")
                return False
    def updates(self, packages: typing.List[str]=[]):
        sub = subprocess.run([self.executable, "-m", "pip", "list", "--outdated", "--format=json"], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        json_str = sub.stdout.decode().strip()
        try:
            tried = json.loads(json_str)
            if packages and len(packages) > 0: return {"success": True, "packages": [i for i in tried if i["name"] in packages]}
            return {"success": True, "packages": tried}
        except: return {"success": False, "packages": []}
    def info(self, packages: typing.List[str]):
        generated_list = [i for i in packages if isinstance(i, str)]
        if len(generated_list) > 0:
            try:
                information = {}
                for i in generated_list:
                    urll = f"https://pypi.org/pypi/{i}/json"
                    if self.getIfConnectedToInternet() == False: return {"success": False}
                    response = self.requests.get(urll)
                    if response.ok:
                        data = response.json
                        info = data["info"]
                        information[i] = info
                return {"success": True, "data": information}
            except Exception as e: return {"success": False}
        return {"success": False}
    def github(self, packages: typing.List[str]):
        generated_list = [i for i in packages if isinstance(i, str)]
        if len(generated_list) > 0:
            try:
                informed = self.info(generated_list)
                if informed["success"] == True:
                    informed = informed["data"]
                    links = {}
                    for i in generated_list:
                        if informed.get(i):
                            info = informed[i]
                            url = info.get("project_urls", {}).get("Source") or info.get("home_page")
                            if url: links[i] = url
                    return {"success": True, "repositories": links}
            except Exception as e: return {"success": False}
        return {"success": False}
    
    # Python Management
    def getLatestPythonVersion(self, beta: bool=False):
        cache_key = f"latest_py_{beta}"
        cached_data = self.__class__._latest_python_cache.get(cache_key)
        if cached_data and (time.time() - cached_data[0]) < 3600: return cached_data[1]
        url = "https://www.python.org/download/pre-releases/" if beta else "https://www.python.org/downloads/"
        response = self.requests.get(url)
        if response.ok: html = response.text
        else: html = ""
        match = re.search(r'Python (\d+\.\d+\.\d+)([a-zA-Z0-9]+)?', html) if beta else re.search(r"Download Python (\d+\.\d+\.\d+)", html)
        if match:
            version = f'{match.group(1)}{match.group(2)}' if beta else match.group(1)
            self.__class__._latest_python_cache[cache_key] = (time.time(), version)
            return version
        self.printDebugMessage("Failed to find latest Python version.")
        return None
    def getLatestPythonInstallManagerVersion(self, beta: bool=False):
        cache_key = f"latest_install_manager_{beta}"
        cached_data = self.__class__._latest_python_cache.get(cache_key)
        if cached_data and (time.time() - cached_data[0]) < 3600: return cached_data[1]
        url = "https://www.python.org/downloads/windows/"
        response = self.requests.get(url)
        if response.ok: html = response.text
        else: html = ""
        match = re.search(r"Python install manager (\d+\.\d+) beta (\d+)", html) if beta else re.search(r"Python install manager (\d+\.\d+)", html)
        if match:
            version = f'{match.group(1)}b{match.group(2)}' if beta else match.group(1)
            self.__class__._latest_python_cache[cache_key] = (time.time(), version)
            return version
        self.printDebugMessage("Failed to find latest Python Installer Manager version.")
        return None
    def getCurrentPythonVersion(self):
        if not self.executable: return None
        if self._current_version_cache is not None: return self._current_version_cache
        if self.isSameRunningPythonExecutable():  self._current_version_cache = platform.python_version()
        else:
            a = subprocess.run([self.executable, "-V"], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            if a.returncode == 0: self._current_version_cache = a.stdout.decode().replace("Python ", "").strip()
        return self._current_version_cache
    def getIfPythonVersionIsBeta(self, version=""):
        cur_vers = version or self.getCurrentPythonVersion()
        match = re.search(r'(\d+\.\d+\.\d+)([a-z]+(\d+)?)?', cur_vers)
        if match:
            _, suf, _ = match.groups()
            return bool(suf)
        return False
    def getIfPythonIsLatest(self):
        cur_vers = self.getCurrentPythonVersion()
        latest_vers = self.getLatestPythonVersion(beta=self.getIfPythonVersionIsBeta())
        return cur_vers == latest_vers
    def pythonInstalled(self, computer=False):
        if computer: return bool(self.findPython())
        if not self.executable: return False
        return os.path.exists(self.executable)
    def extractPythonVersion(self, path, tar_arch: str=None):
        if os.path.islink(path):
            try: target_path = os.path.realpath(path)
            except: target_path = path
        else: target_path = path
        path_lower = target_path.lower() 
        name = os.path.basename(target_path).lower()
        match = re.search(r'python(?:w)?(?:-?|\s*)(\d+)(?:\.(\d+))?(?:\.(\d+))?', name)
        if match: version_tuple = tuple(int(g) if g is not None else 0 for g in match.groups())
        else:
            version_part = os.path.basename(os.path.dirname(target_path))
            match2 = re.match(r'(\d+)(?:\.(\d+))?(?:\.(\d+))?', version_part)
            if match2: version_tuple = tuple(int(g) if g is not None else 0 for g in match2.groups())
            else: version_tuple = (0, 0, 0)
        arch_penalty = 0
        host_arch = os.environ.get("PROCESSOR_ARCHITEW6432", platform.machine()).lower()
        if not host_arch:  host_arch = platform.machine().lower()
        if main_os == "Darwin":
            if host_arch == "arm64" and "intel" in name: arch_penalty = -1
            elif host_arch in ["x86_64", "amd64"] and "arm" in name: arch_penalty = -1
        elif main_os == "Windows":
            is_32bit_path = "-32" in path_lower or "wow64" in path_lower or "(x86)" in path_lower or "\\x86\\" in path_lower
            is_arm_path = "arm64" in path_lower
            if host_arch == "arm64":
                if is_32bit_path: arch_penalty = -2
                elif not is_arm_path:  arch_penalty = -1
            else: 
                if is_32bit_path: arch_penalty = -1 
                elif is_arm_path: arch_penalty = -5 
        return (*version_tuple, arch_penalty)
    def pythonSupported(self, major: int=3, minor: int=13, patch: int=2):
        cur_version = self.getCurrentPythonVersion()
        if not cur_version: return False
        return self.pythonSupportedStatic(cur_version, major, minor, patch)
    def pythonSupportedStatic(self, version: str, major: int=3, minor: int=13, patch: int=2):
        if not version: return False
        match = re.match(r"(\d+)\.(\d+)\.(\w+)", version)
        if match:
            version = match.groups() 
            return tuple(map(self._to_int, version)) >= (major, minor, patch)
        return False
    def osSupported(self, windows_build: int=0, macos_version: tuple=(0,0,0), linux_kernel: tuple=(5,0,0)):
        if main_os == "Windows":
            version = platform.version().split(".")
            if len(version) < 3: return False
            return int(version[2]) >= windows_build
        elif main_os == "Darwin":
            version_tuple = tuple(map(int, platform.mac_ver()[0].split('.')))
            while len(version_tuple) < 3: version_tuple += (0,)
            while len(macos_version) < 3: macos_version += (0,)
            return version_tuple >= macos_version
        elif main_os == "Linux":
            rel = platform.release().split("-")[0]
            if rel:
                version_tuple = tuple(map(self._to_int, rel.split('.')))
                while len(version_tuple) < 3: version_tuple += (0,)
                while len(linux_kernel) < 3: linux_kernel += (0,)
                return version_tuple >= linux_kernel
            return True
        return False
    def pythonInstall(self, version: str="", beta: bool=False, silent: bool=False, manual: bool=False, arch: str=None):
        ma_arch = platform.architecture()
        ma_processor = platform.machine()
        macos_version_numbers = {
            "3.10.0a3": "11.0",
            "3.9.1rc1": "11.0"
        }
        if not self.pythonSupportedStatic(version, 3, 9, 2):
            if self.pythonSupportedStatic(version, 3, 7, 0): macos_version_numbers[version] = "x10.9"
            else: macos_version_numbers[version] = "x10.6"
        if self.getIfConnectedToInternet() == False:
            self.printDebugMessage("Failed to download Python installer.")
            return
        if version == "": version = self.getLatestPythonVersion(beta=beta)
        if not version:
            self.printDebugMessage("Failed to download Python installer.")
            return
        version_url_folder = version
        if beta == True: version_url_folder = re.match(r'^\d+\.\d+\.\d+', version).group()
        if self.uv_exe:
            self.printDebugMessage(f"Using uv to install Python {version}...")
            cmd = [self.uv_exe, "python", "install", version]
            res = subprocess.run(cmd, stdout=None if self.debug else subprocess.DEVNULL, stderr=None if self.debug else subprocess.DEVNULL)
            if res.returncode == 0:
                self.printDebugMessage(f"Successfully installed Python {version} via uv.")
                self.clear_class_caches()
                return
            else: self.printDebugMessage("uv python install failed, falling back to default installation...")
        if main_os in ("Darwin", "Linux") and self.brew_exe:
            major_minor = self.getMajorMinorVersion(version)
            self.printDebugMessage(f"Using Homebrew to install Python {major_minor}...")
            cmd = [self.brew_exe, "install", f"python@{major_minor}"]
            res = subprocess.run(cmd, stdout=None if self.debug else subprocess.DEVNULL, stderr=None if self.debug else subprocess.DEVNULL)
            if res.returncode == 0:
                self.printDebugMessage(f"Successfully installed python@{major_minor} via Homebrew.")
                self.clear_class_caches()
                return
            else: self.printDebugMessage("Homebrew python install failed, falling back to default installation...")
        if main_os == "Linux":
            major_minor = self.getMajorMinorVersion(version)
            for pkg_mgr, install_args in (("apt-get", ["sudo", "apt-get", "install", "-y", f"python{major_minor}"]), ("dnf", ["sudo", "dnf", "install", "-y", f"python{major_minor}"]), ("pacman", ["sudo", "pacman", "-S", "--noconfirm", "python"])):
                if self.getPathFile(pkg_mgr, pkg_mgr):
                    res = subprocess.run(install_args, stdout=None if self.debug else subprocess.DEVNULL, stderr=None if self.debug else subprocess.DEVNULL)
                    if res.returncode == 0:
                        self.clear_class_caches()
                        return
                    break
        if main_os == "Darwin":
            url = f"https://www.python.org/ftp/python/{version_url_folder}/python-{version}-macos{macos_version_numbers.get(version, '11')}.pkg"
            with tempfile.NamedTemporaryFile(suffix=".pkg", delete=False) as temp_file: pkg_file_path = temp_file.name
            result = self.requests.download(url, pkg_file_path)            
            if result.ok:
                if silent == True: 
                    self.printDebugMessage(f"Silently installing Python packages.. Admin permissions may be requested.")
                    subprocess.run(["osascript", "-e", f"do shell script \"installer -pkg '{pkg_file_path}' -target /\" with administrator privileges"], stdout=self.debug == False and subprocess.DEVNULL, stderr=self.debug == False and subprocess.DEVNULL, check=True)
                    self.printDebugMessage(f"Successfully installed Python package: {pkg_file_path}")
                else:
                    subprocess.run(["open", pkg_file_path], stdout=self.debug == False and subprocess.DEVNULL, stderr=self.debug == False and subprocess.DEVNULL, check=True)
                    while self.getIfProcessIsOpened("/System/Library/CoreServices/Installer.app") == True: time.sleep(0.1)
                    self.printDebugMessage(f"Python installer has been executed: {pkg_file_path}")
                self.clear_class_caches()
            else: self.printDebugMessage("Failed to download Python installer.")
        elif main_os == "Windows":
            if not self.pythonSupportedStatic(version, 3, 11, 0): self.printDebugMessage("PyKits is not normally made for versions less than 3.11.0.")
            if arch == "arm64":
                ma_processor = "arm64"
                ma_arch = ["64bit"]
            elif arch == "x64":
                ma_processor = "amd64"
                ma_arch = ["64bit"]
            elif arch == "x86":
                ma_arch = ["32bit"]
            if (manual == True and self.pythonSupportedStatic(version, 3, 5, 0)) or self.pythonSupportedStatic(version, 3, 15, 0):
                self.printDebugMessage("Setting up python modules..")
                if main_os == "Windows" and (not hasattr(self, "_win32api") or not hasattr(self, "_win32con") or not hasattr(self, "_win32gui")):
                    try:
                        try:
                            import win32api # type: ignore
                            import win32con # type: ignore
                            import win32gui # type: ignore
                            self._win32con = win32con
                            self._win32api = win32api
                            self._win32gui = win32gui
                        except Exception:
                            self.install(["pywin32"])
                            self._win32con = self.importModule("win32con")
                            self._win32api = self.importModule("win32api")
                            self._win32gui = self.importModule("win32gui")
                    except: pass
                self.printDebugMessage("Downloading Python package..")
                arch_fold_end = ""
                if ma_arch[0] == "64bit":
                    if ma_processor.lower() == "arm64": url = f"https://www.python.org/ftp/python/{version_url_folder}/python-{version}-arm64.zip"; arch_fold_end = "-arm64"
                    else: url = f"https://www.python.org/ftp/python/{version_url_folder}/python-{version}-amd64.zip"
                else: url = f"https://www.python.org/ftp/python/{version_url_folder}/python-{version}-win32.zip"; arch_fold_end = "-32"
                with tempfile.NamedTemporaryFile(suffix=".zip", delete=False) as temp_file: zip_file_path = temp_file.name
                result = self.requests.download(url, zip_file_path)
                if result.ok:
                    stripped_version = "".join(self.getMajorMinorVersion(version_url_folder).split("."))
                    self.printDebugMessage(f"Unzipping Python {version_url_folder} package..")
                    user_appdata_folder = self.getLocalAppData()
                    target_python_path = os.path.join(user_appdata_folder, "Programs", "Python", f"Python{stripped_version}{arch_fold_end}")
                    unzip_res = self.unzipFile(zip_file_path, target_python_path, ["python.exe"])
                    if unzip_res.returncode == 0:
                        self.printDebugMessage(f"Installing PIP..")
                        pip_script_path = os.path.join(target_python_path, "get-pip.py")
                        if self.requests.download("https://bootstrap.pypa.io/get-pip.py", pip_script_path).ok:
                            python_exe = os.path.join(target_python_path, "python.exe")
                            install_pip_res = subprocess.run([python_exe, pip_script_path], cwd=target_python_path)
                            if install_pip_res.returncode == 0: self.printDebugMessage("Successfully installed PIP with Python!")
                            else: self.printDebugMessage(f"Unable to install PIP with Python Package. Return Code: {install_pip_res.returncode}")
                            os.remove(pip_script_path)
                            if (not shutil.which("python")) or self.getArchitecture() == pip(executable=python_exe).getArchitecture():
                                self.printDebugMessage("Adding to user path..")
                                self.addToUserPath(target_python_path)
                                self.addToUserPath(os.path.join(target_python_path, "Scripts"))
                                self.printDebugMessage("Registering file extensions..")
                                extensions = [(".py", "Python.File"), (".pyw", "Python.NoConFile")]
                                for ext, prog_id in extensions:
                                    ext_key = self._win32api.RegCreateKey(self._win32con.HKEY_CURRENT_USER, f"Software\\Classes\\{ext}")
                                    self._win32api.RegSetValueEx(ext_key, "", 0, self._win32con.REG_SZ, prog_id)
                                    self._win32api.RegCloseKey(ext_key)
                                    cmd_key = self._win32api.RegCreateKey(self._win32con.HKEY_CURRENT_USER, f"Software\\Classes\\{prog_id}\\shell\\open\\command")
                                    self._win32api.RegSetValueEx(cmd_key, "", 0, self._win32con.REG_SZ, f'"{python_exe}" "%1" %*')
                                    self._win32api.RegCloseKey(cmd_key)
                                self._win32gui.SendMessageTimeout(
                                    self._win32con.HWND_BROADCAST,
                                    self._win32con.WM_SETTINGCHANGE,
                                    0,
                                    "Environment",
                                    self._win32con.SMTO_ABORTIFHUNG,
                                    5000
                                )
                            registry, shortcuts = self.registerPython(target_python_path, version)
                            if registry: self.printDebugMessage(f"Successfully installed Python {version} into path: {target_python_path}")
                            else: self.printDebugMessage("Failed to add Python to the registry.")
                        else: self.printDebugMessage("Failed to bootstrap pip (download get-pip.py failed).")
                        self.clear_class_caches()
                    else: self.printDebugMessage("Failed to download Python installer.")
                else: self.printDebugMessage("Failed to download Python installer.")
            else:
                if ma_arch[0] == "64bit":
                    if ma_processor.lower() == "arm64": url = f"https://www.python.org/ftp/python/{version_url_folder}/python-{version}-arm64.exe"
                    else: url = f"https://www.python.org/ftp/python/{version_url_folder}/python-{version}-amd64.exe"
                else: url = f"https://www.python.org/ftp/python/{version_url_folder}/python-{version}.exe"
                with tempfile.NamedTemporaryFile(suffix=".exe", delete=False) as temp_file: exe_file_path = temp_file.name
                result = self.requests.download(url, exe_file_path)
                if result.ok:
                    if silent == True:
                        self.printDebugMessage(f"Silently installing Python packages..")
                        subprocess.run([exe_file_path, "/quiet"], stdout=self.debug == False and subprocess.DEVNULL, stderr=self.debug == False and subprocess.DEVNULL, check=True)
                        self.printDebugMessage(f"Successfully installed Python package: {exe_file_path}")
                    else:
                        subprocess.run([exe_file_path], stdout=self.debug == False and subprocess.DEVNULL, stderr=self.debug == False and subprocess.DEVNULL, check=True)
                        self.printDebugMessage(f"Python installer has been executed: {exe_file_path}")
                    self.clear_class_caches()
                else: self.printDebugMessage("Failed to download Python installer.")
    def registerPython(self, target_python_path: str, version: str):
        """Only works on Windows."""
        if main_os != "Windows": return
        major_minor = self.getMajorMinorVersion(version)
        python_exe = os.path.join(target_python_path, "python.exe")
        pythonw_exe = os.path.join(target_python_path, "pythonw.exe")
        passed1 = True
        passed2 = True
        try:
            core_key_path = rf"Software\Python\PythonCore\{major_minor}"
            base_key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, core_key_path)
            winreg.SetValueEx(base_key, "DisplayName", 0, winreg.REG_SZ, f"Python {major_minor}")
            winreg.SetValueEx(base_key, "SupportUrl", 0, winreg.REG_SZ, "https://www.python.org/")
            winreg.SetValueEx(base_key, "Version", 0, winreg.REG_SZ, version)
            install_key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, rf"{core_key_path}\InstallPath")
            winreg.SetValue(install_key, "", winreg.REG_SZ, target_python_path)
            winreg.SetValueEx(install_key, "ExecutablePath", 0, winreg.REG_SZ, python_exe)
            winreg.SetValueEx(install_key, "WindowedExecutablePath", 0, winreg.REG_SZ, pythonw_exe)
            path_key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, rf"{core_key_path}\PythonPath")
            lib_path = os.path.join(target_python_path, "Lib")
            dlls_path = os.path.join(target_python_path, "DLLs")
            winreg.SetValue(path_key, "", winreg.REG_SZ, f"{lib_path};{dlls_path}")
            winreg.CloseKey(path_key)
            winreg.CloseKey(install_key)
            winreg.CloseKey(base_key)
            self.printDebugMessage(f"Successfully registered Python {version} in the Registry.")
        except Exception as e: self.printDebugMessage(f"Failed to write to registry: {e}"); passed1 = False
        try:
            win32com = self.importModule("win32com.client", install_module_if_not_found=True)
            shell = win32com.Dispatch("WScript.Shell")
            major_minor = self.getMajorMinorVersion(version)
            python_exe = os.path.join(target_python_path, "python.exe")
            start_menu = shell.SpecialFolders("Programs")
            shortcut_folder = os.path.join(start_menu, f"Python {major_minor}")
            if not os.path.exists(shortcut_folder): os.makedirs(shortcut_folder); os.chmod(shortcut_folder, 0o755)
            cli_shortcut_path = os.path.join(shortcut_folder, f"Python {major_minor}.lnk")
            cli_shortcut = shell.CreateShortCut(cli_shortcut_path)
            cli_shortcut.Targetpath = python_exe
            cli_shortcut.WorkingDirectory = target_python_path
            cli_shortcut.IconLocation = f"{python_exe},0"
            cli_shortcut.Description = f"Python {major_minor} CLI"
            cli_shortcut.save()
            idle_path = os.path.join(target_python_path, "Lib", "idlelib", "idle.pyw")
            if os.path.exists(idle_path):
                pythonw_exe = os.path.join(target_python_path, "pythonw.exe")
                idle_shortcut_path = os.path.join(shortcut_folder, f"IDLE (Python {major_minor}).lnk")
                idle_shortcut = shell.CreateShortCut(idle_shortcut_path)
                idle_shortcut.Targetpath = pythonw_exe
                idle_shortcut.Arguments = f'"{idle_path}"'
                idle_shortcut.WorkingDirectory = target_python_path
                idle_shortcut.IconLocation = f"{idle_path},0"
                idle_shortcut.Description = f"IDLE (Python {major_minor})"
                idle_shortcut.save()
            self.printDebugMessage(f"Successfully created Start Menu shortcuts for Python {major_minor}.")
        except Exception as e: self.printDebugMessage(f"Failed to create shortcuts: {e}"); passed2 = False
        return (passed1, passed2)
    def findPythonInstallManager(self): return shutil.which("python-install") or shutil.which("python-install-manager")
    def installLocalPythonCertificates(self):
        if main_os == "Darwin":
            with open("./install_local_python_certs.py", "w") as f: f.write("""import os; import os.path; import ssl; import stat; import subprocess; import sys; STAT_0o775 = ( stat.S_IRUSR | stat.S_IWUSR | stat.S_IXUSR | stat.S_IRGRP | stat.S_IWGRP | stat.S_IXGRP | stat.S_IROTH |  stat.S_IXOTH ); openssl_dir, openssl_cafile = os.path.split(ssl.get_default_verify_paths().openssl_cafile); print(" -- pip install --upgrade certifi"); subprocess.check_call([sys.executable, "-E", "-s", "-m", "pip", "install", "--upgrade", "certifi"]); import certifi; os.chdir(openssl_dir); relpath_to_certifi_cafile = os.path.relpath(certifi.where()); print(" -- removing any existing file or link"); os.remove(openssl_cafile); print(" -- creating symlink to certifi certificate bundle"); os.symlink(relpath_to_certifi_cafile, openssl_cafile); print(" -- setting permissions"); os.chmod(openssl_cafile, STAT_0o775); print(" -- update complete");""")
            s = subprocess.run(f'"{self.executable}" ./install_local_python_certs.py', shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            os.remove("./install_local_python_certs.py")
            if s.returncode != 0 and self.debug == True: print(f"Unable to install local python certificates!")
    def getIf32BitWindows(self): return main_os == "Windows" and self.getArchitecture() == "x86"
    def getIfArmWindows(self): return main_os == "Windows" and self.getArchitecture() == "arm"
    def getIfRunningWindowsAdmin(self):
        if main_os == "Windows":
            try: return ctypes.windll.shell32.IsUserAnAdmin()
            except: return False
        return False
    def getArchitecture(self):
        if self._arch_cache is not None: return self._arch_cache
        exe = self.executable if self.executable else sys.executable
        if self.isSameRunningPythonExecutable():
            machine_var = platform.machine()
            if main_os == "Windows":
                with open(exe, "rb") as f:
                    mm = mmap.mmap(f.fileno(), 0, access=mmap.ACCESS_READ)
                    pe_offset = int.from_bytes(mm[0x3C:0x40], "little")
                    machine = int.from_bytes(mm[pe_offset + 4:pe_offset + 6], "little")
                    mm.close()
                arch_map = { 0x014c: "x86", 0x8664: "x64", 0xAA64: "arm", 0x01c0: "arm" }
                self._arch_cache = arch_map.get(machine, "")
            elif main_os == "Darwin": self._arch_cache = "arm" if machine_var.lower() == "arm64" else ("intel" if machine_var.lower() == "x86_64" else "x86")
            else: self._arch_cache = machine_var
        else:
            if main_os == "Darwin":
                try:
                    exe_name = os.path.basename(exe).lower()
                    if "intel" in exe_name:
                        self._arch_cache = "intel"
                        return self._arch_cache
                    with open(exe, "rb") as f:
                        header = f.read(68)
                    if len(header) >= 8:
                        magic = int.from_bytes(header[:4], "big")
                        if magic in (0xFEEDFACF, 0xFEEDFACE):
                            cputype = int.from_bytes(header[4:8], "big")
                            self._arch_cache = "arm" if cputype == 0x0100000C else ("intel" if cputype == 0x01000007 else ("x86" if cputype == 7 else ""))
                        elif magic in (0xCFFAEDFE, 0xCEFAEDFE):
                            cputype = int.from_bytes(header[4:8], "little")
                            self._arch_cache = "arm" if cputype == 0x0100000C else ("intel" if cputype == 0x01000007 else ("x86" if cputype == 7 else ""))
                        elif magic in (0xCAFEBABE, 0xCAFEBABF):
                            nfat = int.from_bytes(header[4:8], "big")
                            cputypes = set()
                            stride = 20 if magic == 0xCAFEBABE else 32
                            for ind in range(min(nfat, 3)):
                                off = 8 + ind * stride
                                if len(header) >= off + 4:
                                    cputypes.add(int.from_bytes(header[off:off+4], "big"))
                            host_is_arm = platform.machine().lower() == "arm64"
                            if host_is_arm and 0x0100000C in cputypes: self._arch_cache = "arm"
                            elif 0x01000007 in cputypes: self._arch_cache = "intel"
                            elif 0x0100000C in cputypes: self._arch_cache = "arm"
                            elif 7 in cputypes: self._arch_cache = "x86"
                    if not self._arch_cache:
                        s = subprocess.run([exe, "-c", "import platform; machine_var = platform.machine(); print('arm' if machine_var.lower() == 'arm64' else ('intel' if machine_var.lower() == 'x86_64' else 'x86'))"], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                        self._arch_cache = s.stdout.decode().strip()
                except Exception: self._arch_cache = ""
            elif main_os == "Windows":
                try:
                    with open(exe, "rb") as f:
                        mm = mmap.mmap(f.fileno(), 0, access=mmap.ACCESS_READ)
                        pe_offset = int.from_bytes(mm[0x3C:0x40], "little")
                        machine = int.from_bytes(mm[pe_offset + 4:pe_offset + 6], "little")
                        mm.close()
                    arch_map = { 0x014c: "x86", 0x8664: "x64", 0xAA64: "arm", 0x01c0: "arm" }
                    self._arch_cache = arch_map.get(machine, "")
                except: self._arch_cache = ""
            else:
                result = subprocess.run([exe, "-c", "import platform; print(platform.machine())"], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                self._arch_cache = result.stdout.decode().strip()
        return self._arch_cache
    def getIfVirtualEnvironment(self):
        alleged_path = os.path.dirname(self.executable)
        return os.path.exists(os.path.join(alleged_path, "..", "pyvenv.cfg")) or (os.path.exists(os.path.join(alleged_path, "python.exe")) and os.path.exists(os.path.join(alleged_path, "pip.exe")))
    def findPython(self, arch=None, latest=True, optimize=True, path=False):
        if main_os == "Darwin":
            target_name = "python3-intel64" if arch == "intel" else "python3"
            target_path = f"/usr/local/bin/{target_name}"
            if optimize == True and os.path.exists(target_path) and os.path.islink(target_path): return target_path if path == True else pip(executable=target_path, find=False)
        all_pythons = self.findPythons(arch=arch, latest=latest, paths=path)
        if len(all_pythons) > 0: return all_pythons[0]
        return None
    def findPythons(self, arch=None, latest=True, paths=False):
        cache_key = f"pythons_{arch}_{latest}_{paths}"
        cached_data = self.__class__._python_path_cache.get(cache_key)
        if cached_data and (time.time() - cached_data[0]) < 60: return cached_data[1]
        raw_paths = set()
        path_table = []
        if main_os == "Darwin":
            path_table = [
                "/usr/local/bin/python*", "/opt/homebrew/bin/python*",
                "/Library/Frameworks/Python.framework/Versions/*/bin/python*",
                os.path.expanduser("~/Library/Python/*/bin/python*"),
                os.path.expanduser("~/.pyenv/versions/*/bin/python*"),
                os.path.expanduser("~/opt/anaconda*/bin/python*")
            ]
            if self.brew_exe:
                try:
                    brew_prefix = subprocess.check_output([self.brew_exe, "--prefix"], text=True).strip()
                    path_table.append(f"{brew_prefix}/bin/python*")
                except Exception: pass
        elif main_os == "Linux":
            path_table = [
                "/usr/bin/python*", "/usr/local/bin/python*",
                os.path.expanduser("~/.local/bin/python*"),
                os.path.expanduser("~/.pyenv/versions/*/bin/python*"),
                os.path.expanduser("~/opt/anaconda*/bin/python*")
            ]
            if self.brew_exe:
                try:
                    brew_prefix = subprocess.check_output([self.brew_exe, "--prefix"], text=True).strip()
                    path_table.append(f"{brew_prefix}/bin/python*")
                except Exception: pass
        elif main_os == "Windows":
            path_table = [
                os.path.expandvars(r'%LOCALAPPDATA%\Programs\Python\Python*\python.exe'),
                os.path.expandvars(r'%PROGRAMFILES%\Python*\python.exe'),
                os.path.expandvars(r'%PROGRAMFILES(x86)%\Python*\python.exe'),
                os.path.join(os.environ.get("RUNNER_TOOL_CACHE", r"C:\hostedtoolcache\windows"), "Python", "*", "*", "python.exe")
            ]
        if self.uv_exe:
            try:
                uv_python_dir = subprocess.check_output([self.uv_exe, "python", "dir"], text=True).strip()
                if main_os == "Windows":
                    path_table.append(os.path.join(uv_python_dir, "*", "python.exe"))
                    path_table.append(os.path.join(uv_python_dir, "*", "install", "python.exe"))
                else:
                    path_table.append(os.path.join(uv_python_dir, "*", "bin", "python*"))
                    path_table.append(os.path.join(uv_python_dir, "*", "install", "bin", "python*"))
            except Exception: pass
        for pattern in path_table: raw_paths.update(glob.glob(pattern))
        valid_paths = []
        for p in raw_paths:
            if os.path.isfile(p):
                basename = os.path.basename(p).lower()
                if main_os == "Windows" and ".exe" in basename: basename = basename.replace(".exe", "")
                if basename.startswith("pythonw") or basename.endswith("config") or basename.endswith("m") or basename.endswith("t") or basename.endswith("remotezip2"): continue
                if "WindowsApps" in p: continue
                valid_paths.append(p)
        if latest == True: valid_paths.sort(reverse=True, key=self.extractPythonVersion)
        founded_pythons = []
        for p in valid_paths:
            pip_class = pip(executable=p, find=False) 
            if arch:
                py_arch = pip_class.getArchitecture()
                if py_arch == "": continue
                if py_arch == arch: founded_pythons.append(p if paths == True else pip_class)
            else: founded_pythons.append(p if paths == True else pip_class)
        self.__class__._python_path_cache[cache_key] = (time.time(), founded_pythons)
        return founded_pythons
    def findPythonViaRegistry(self):
        if main_os != "Windows": return []
        found_paths = []
        hkeys = [winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE]
        access_flags = [
            winreg.KEY_READ | winreg.KEY_WOW64_64KEY,  # 64
            winreg.KEY_READ | winreg.KEY_WOW64_32KEY   # 32
        ]
        for hkey in hkeys:
            for flag in access_flags:
                try:
                    core_key = winreg.OpenKey(hkey, r"Software\Python\PythonCore", 0, flag)
                    num_of_subkeys = winreg.QueryInfoKey(core_key)[0]
                    for i in range(num_of_subkeys):
                        version_str = winreg.EnumKey(core_key, i)
                        try:
                            install_key = winreg.OpenKey(core_key, rf"{version_str}\InstallPath", 0, flag)
                            executable_path, _ = winreg.QueryValueEx(install_key, "ExecutablePath")
                            if os.path.exists(executable_path) and executable_path not in found_paths: found_paths.append(executable_path)
                            winreg.CloseKey(install_key)
                        except OSError: continue
                    winreg.CloseKey(core_key)
                except OSError: continue
        return found_paths
    def isSameRunningPythonExecutable(self):
        if self.ignore_same: return False
        if not self.executable: return False
        if self._is_same_env is not None: return self._is_same_env
        if os.path.exists(self.executable) and os.path.exists(sys.executable): self._is_same_env = os.path.samefile(self.executable, sys.executable)
        else: self._is_same_env = False
        return self._is_same_env
    def getMajorMinorVersion(self, version: str="3.14.7"): return ".".join(version.split(".")[:-1])

    # Python Functions
    def getLocalAppData(self):
        if main_os == "Windows": return os.path.expandvars(r'%LOCALAPPDATA%')
        elif main_os == "Darwin": return f'{os.path.expanduser("~")}/Library/'
        elif main_os == "Linux": return os.environ.get("XDG_DATA_HOME") or os.path.join(os.path.expanduser("~"), ".local", "share")
        else: return f'{os.path.expanduser("~")}/'
    def getUserFolder(self): return os.path.expanduser("~")
    def getIfLoggedInIsMacOSAdmin(self):
        if main_os == "Darwin":
            logged_in_folder = self.getUserFolder()
            username = os.path.basename(logged_in_folder)
            groups_res = subprocess.run([self.getPathFile("/usr/bin/groups"), username], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            if groups_res.returncode == 0: return "admin" in groups_res.stdout.decode("utf-8").split(" ")
            else: return False
        else: return False
    def getInstallableApplicationsFolder(self):
        if main_os == "Darwin":
            if self.getIfLoggedInIsMacOSAdmin(): return os.path.join("/", "Applications")
            else: return os.path.join(self.getUserFolder(), "Applications")
        elif main_os in ("Windows", "Linux"):
            return self.getLocalAppData()
    def restartScript(self, scriptname: str, argv: list):
        argv.pop(0)
        res = subprocess.run([self.executable, os.path.join(os.path.dirname(os.path.abspath(__file__)), scriptname)] + argv)
        sys.exit(res.returncode)
    def clearStdin(self):
        if os.name == 'nt': 
            while msvcrt.kbhit(): msvcrt.getch()
        else: termios.tcflush(sys.stdin, termios.TCIFLUSH)
    def endProcess(self, name="", pid=""):
        if pid == "":
            if main_os == "Darwin": subprocess.run([self.getPathFile("/usr/bin/killall"), "-9", name], stdout=subprocess.DEVNULL)
            elif main_os == "Windows": subprocess.run(f"taskkill /IM {name} /F", shell=True, stdout=subprocess.DEVNULL)
            else: subprocess.run(f"killall -9 {name}", shell=True, stdout=subprocess.DEVNULL)
        else:
            if main_os == "Darwin": subprocess.run(f"kill -9 {pid}", shell=True, stdout=subprocess.DEVNULL)
            elif main_os == "Windows": subprocess.run(f"taskkill /PID {pid} /F", shell=True, stdout=subprocess.DEVNULL)
            else: subprocess.run(f"kill -9 {pid}", shell=True, stdout=subprocess.DEVNULL)
    def importModule(self, module_name: str, install_module_if_not_found: bool=False, loop_until_import: bool=False):
        existing = sys.modules.get(module_name)
        if existing is not None: return existing
        try: 
            s = importlib.import_module(module_name)
            if s is None: raise ModuleNotFoundError("")
            else: return s
        except ModuleNotFoundError:
            try:
                if install_module_if_not_found == True and self.isSameRunningPythonExecutable(): self.install([module_name])
                self.uncacheLoadedModules()
                s = importlib.import_module(module_name)
                if s is None: raise ModuleNotFoundError("")
                else: return s
            except Exception: 
                if loop_until_import == False: raise ImportError(f'Unable to find module "{module_name}" in Python {self.getCurrentPythonVersion()} environment.')
                time.sleep(1)
                return self.importModule(module_name=module_name, install_module_if_not_found=install_module_if_not_found, loop_until_import=loop_until_import)
        except Exception as e: raise ImportError(f'Unable to import module "{module_name}" in Python {self.getCurrentPythonVersion()} environment. Exception: {str(e)}')
    def importModules(self, modules: typing.List[str], install_modules_if_not_found: bool=False):
        modules_collected = []
        if install_modules_if_not_found == True:
            installed = self.installed(modules)
            needing_install = []
            for i, v in installed.items():
                if i != "all" and v == False: needing_install.append(i)
            if len(needing_install) > 0:
                self.install(needing_install)
                self.uncacheLoadedModules()
        for module_name in modules:
            existing = sys.modules.get(module_name)
            if existing is not None:
                modules_collected.append(existing)
                continue
            try: 
                s = importlib.import_module(module_name)
                if s is None: raise ModuleNotFoundError("")
                modules_collected.append(s)
            except ModuleNotFoundError:
                try:
                    self.uncacheLoadedModules()
                    s = importlib.import_module(module_name)
                    if s is None: raise ModuleNotFoundError("")
                    modules_collected.append(s)
                except Exception as e: raise ImportError(f'Unable to import module "{module_name}" in Python {self.getCurrentPythonVersion()} environment. Exception: {str(e)}')
            except Exception as e: raise ImportError(f'Unable to import module "{module_name}" in Python {self.getCurrentPythonVersion()} environment. Exception: {str(e)}')
        return tuple(modules_collected)
    def uncacheLoadedModules(self):
        if getattr(sys, "frozen", False) or "__compiled__" in globals(): pass
        else:
            if self.isSameRunningPythonExecutable() and self.getIfVirtualEnvironment(): return
            import site
            self._site = site
            site_packages_paths = self._site.getsitepackages() + [self._site.getusersitepackages()]
            for path in site_packages_paths:
                if path not in sys.path and os.path.exists(path): sys.path.append(path)
        importlib.invalidate_caches()
    def unzipFile(self, path: str, output: str, look_for: typing.List=None, export_out: typing.List=None, either: bool=False, check: bool=True, moving_file_func: typing.Callable=None):
        class Result():
            __slots__ = ("ok", "returncode", "path")
            def __init__(self):
                self.ok = True
                self.returncode = 0
                self.path = ""
        if not look_for: look_for = []
        if not export_out: export_out = []
        if not os.path.exists(output): os.makedirs(output); os.chmod(output, 0o755)
        previous_output = output
        if output.endswith("/"): output = output[:-1]
        tar_exe = self.getPathFile("C:\\Windows\\System32\\tar.exe") if main_os == "Windows" else (self.getPathFile("/usr/bin/unzip", "unzip") if main_os == "Linux" else self.getPathFile("/usr/bin/ditto"))
        if not look_for:
            if main_os == "Windows": subprocess.run([tar_exe, "-xf", path] + export_out + ["-C", output], stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=check)
            elif main_os == "Linux":
                if tar_exe:
                    proc = subprocess.run([tar_exe, "-q", "-o", path] + export_out + ["-d", output], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                    if proc.returncode not in (0, 1, 2):
                        try:
                            with zipfile.ZipFile(path, "r") as zf: zf.extractall(output)
                        except Exception:
                            if check: raise subprocess.CalledProcessError(proc.returncode, proc.args, proc.stdout, proc.stderr)
                else:
                    with zipfile.ZipFile(path, "r") as zf: zf.extractall(output)
            else: subprocess.run([tar_exe, "-xk", path, output], stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=check)
            s = Result()
            s.path = previous_output
            s.returncode = 0
            s.ok = True
            return s
        with tempfile.TemporaryDirectory(prefix="PyKitsUnzip_") as temp_dir:
            if main_os == "Windows": zip_extract = subprocess.run([tar_exe, "-xf", path] + export_out + ["-C", temp_dir], stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=check)
            elif main_os == "Linux":
                if tar_exe:
                    zip_extract = subprocess.run([tar_exe, "-q", "-o", path] + export_out + ["-d", temp_dir], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                    if zip_extract.returncode in (1, 2):
                        zip_extract = Result()
                    elif zip_extract.returncode != 0:
                        try:
                            with zipfile.ZipFile(path, "r") as zf: zf.extractall(temp_dir)
                            zip_extract = Result()
                        except Exception:
                            if check: raise subprocess.CalledProcessError(zip_extract.returncode, zip_extract.args, zip_extract.stdout, zip_extract.stderr)
                else:
                    with zipfile.ZipFile(path, "r") as zf: zf.extractall(temp_dir)
                    zip_extract = Result()
            else: zip_extract = subprocess.run([tar_exe, "-xk", path, temp_dir], stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=check)
            if zip_extract.returncode == 0:
                look_set = set(look_for)
                for ro, dir, fi in os.walk(temp_dir):
                    l = set(fi) | set(dir)
                    found_all = bool(look_set & l) if either else look_set.issubset(l)
                    if found_all: 
                        if moving_file_func: moving_file_func()
                        if os.path.exists(previous_output): shutil.rmtree(previous_output, ignore_errors=True)
                        shutil.move(ro, previous_output)
                        s = Result()
                        s.path = previous_output
                        s.returncode = 0
                        s.ok = True
                        return s
        if os.path.exists(previous_output): shutil.rmtree(previous_output, ignore_errors=True)
        s = Result()
        s.path = None
        s.returncode = 1
        s.ok = False
        return s
    def getPathFile(self, file: str, name: str=""):
        if os.path.exists(file): return file
        elif not name: return shutil.which(os.path.basename(file).replace(".exe", "")) if os.path.basename(file).endswith(".exe") else shutil.which(os.path.basename(file))
        else: return shutil.which(name)
    def copyTreeWithMetadata(self, src: str, dst: str, symlinks=False, ignore=None, dirs_exist_ok=False, ignore_if_not_exist=False):
        if not os.path.exists(src) and ignore_if_not_exist == False: return
        if not dirs_exist_ok and os.path.exists(dst): raise FileExistsError(f"Destination '{dst}' already exists.")
        os.makedirs(dst, exist_ok=True)
        os.chmod(dst, 0o755)
        for root, dirs, files in os.walk(src):
            rel_path = os.path.relpath(root, src)
            dst_root = os.path.join(dst, rel_path)
            ignored_names = (ignore(root, os.listdir(root)) if callable(ignore) else set(ignore)) if ignore else set()
            dirs[:] = [d for d in dirs if d not in ignored_names]
            files = [f for f in files if f not in ignored_names]
            os.makedirs(dst_root, exist_ok=True)
            os.chmod(dst_root, 0o755)
            for dir_name in dirs:
                src_dir = os.path.join(root, dir_name)
                dst_dir = os.path.join(dst_root, dir_name)
                if os.path.islink(src_dir) and symlinks:
                    link_target = os.readlink(src_dir)
                    os.symlink(link_target, dst_dir)
                else:
                    os.makedirs(dst_dir, exist_ok=True)
                    os.chmod(dst_dir, 0o755)
            for file_name in files:
                src_file = os.path.join(root, file_name)
                dst_file = os.path.join(dst_root, file_name)
                if os.path.islink(src_file) and symlinks:
                    link_target = os.readlink(src_file)
                    os.symlink(link_target, dst_file)
                else:
                    shutil.copy2(src_file, dst_file)
                    src_mode = os.stat(src_file).st_mode
                    is_executable = bool(src_mode & stat.S_IXUSR)
                    if is_executable: os.chmod(dst_file, 0o755)
                    else: os.chmod(dst_file, 0o644)
            shutil.copystat(root, dst_root, follow_symlinks=False)
            os.chmod(dst_root, 0o755)
        return dst
    def _snapshot_windows_processes(self):
        try:
            TH32CS_SNAPPROCESS = 0x00000002
            INVALID_HANDLE_VALUE = ctypes.c_void_p(-1).value
            class PROCESSENTRY32W(ctypes.Structure):
                _fields_ = [
                    ("dwSize", wintypes.DWORD),
                    ("cntUsage", wintypes.DWORD),
                    ("th32ProcessID", wintypes.DWORD),
                    ("th32DefaultHeapID", ctypes.c_size_t),
                    ("th32ModuleID", wintypes.DWORD),
                    ("cntThreads", wintypes.DWORD),
                    ("th32ParentProcessID", wintypes.DWORD),
                    ("pcPriClassBase", wintypes.LONG),
                    ("dwFlags", wintypes.DWORD),
                    ("szExeFile", wintypes.WCHAR * 260),
                ]
            kernel32 = ctypes.windll.kernel32
            snap = kernel32.CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
            if snap == INVALID_HANDLE_VALUE or not snap: return None
            try:
                pe = PROCESSENTRY32W()
                pe.dwSize = ctypes.sizeof(PROCESSENTRY32W)
                procs = []
                if kernel32.Process32FirstW(snap, ctypes.byref(pe)):
                    while True:
                        procs.append((pe.th32ProcessID, pe.szExeFile))
                        if not kernel32.Process32NextW(snap, ctypes.byref(pe)): break
                return procs
            finally: kernel32.CloseHandle(snap)
        except Exception: return None
    def getIfProcessIsOpened(self, process_name="", pid=""):
        if main_os == "Windows":
            procs = self._snapshot_windows_processes()
            if procs is not None:
                if pid: return any(str(p) == str(pid) for p, _ in procs)
                if process_name: return any(process_name.lower() in name.lower() for _, name in procs)
                return False
            no_win = getattr(subprocess, "CREATE_NO_WINDOW", 0)
            process_list = subprocess.run(["tasklist"], stdout=subprocess.PIPE, stderr=subprocess.PIPE, creationflags=no_win).stdout.decode("utf-8", errors="ignore")
            if not pid: return process_name in process_list
            else: return any(line.split()[1] == str(pid) for line in process_list.splitlines() if len(line.split()) > 1)
        else:
            if not pid: return subprocess.run(f"pgrep -f '{process_name}' > /dev/null 2>&1", shell=True).returncode == 0
            else: return subprocess.run(f"ps -p {pid} > /dev/null 2>&1", shell=True).returncode == 0
    def getAmountOfProcesses(self, process_name=""):
        if main_os == "Windows":
            procs = self._snapshot_windows_processes()
            if procs is not None: return sum(1 for _, name in procs if process_name.lower() in name.lower())
            no_win = getattr(subprocess, "CREATE_NO_WINDOW", 0)
            process = subprocess.Popen(["tasklist"], stdout=subprocess.PIPE, stderr=subprocess.PIPE, creationflags=no_win)
            output, _ = process.communicate()
            process_list = output.decode("utf-8", errors="ignore")
            return process_list.lower().count(process_name.lower())
        else:
            result = subprocess.run(f"pgrep -f '{process_name}'", stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)
            process_ids = result.stdout.decode("utf-8").strip().split("\n")
            return len([pid for pid in process_ids if pid.isdigit()])
    def getIfConnectedToInternet(self): return self.requests.get_if_connected()
    def startThread(self, func: typing.Callable, daemon: bool=False, *args, **kwargs):
        thread = threading.Thread(target=func, args=args, kwargs=kwargs, daemon=daemon)
        thread.start()
        self._daemon_threads.add(thread)
        return thread
    def delayedThread(self, func: typing.Callable, time: float=3, daemon: bool=False, *args, **kwargs):
        thread = threading.Timer(time, func, args=args, kwargs=kwargs)
        thread.daemon = daemon
        thread.start()
        self._daemon_threads.add(thread)
        return thread
    def getProcessWindows(self, pid: int):
        if main_os == "Windows" and (not hasattr(self, "_win32gui") or not hasattr(self, "_win32process")):
            try:
                try:
                    import win32gui # type: ignore
                    import win32process # type: ignore
                    self._win32gui = win32gui
                    self._win32process = win32process
                except Exception:
                    self.install(["pywin32"])
                    self._win32gui = self.importModule("win32gui")
                    self._win32process = self.importModule("win32process")
            except: pass
        elif main_os == "Darwin" and (not hasattr(self, "_CGWindowListCopyWindowInfo") or not hasattr(self, "_kCGWindowListOptionOnScreenOnly")):
            try:
                try:
                    from Quartz import CGWindowListCopyWindowInfo, kCGWindowListOptionOnScreenOnly # type: ignore
                except Exception as e:
                    self.install(["pyobjc-framework-Quartz"])
                    Quartz = self.importModule("Quartz")
                    CGWindowListCopyWindowInfo, kCGWindowListOptionOnScreenOnly = Quartz.CGWindowListCopyWindowInfo, Quartz.kCGWindowListOptionOnScreenOnly
                self._CGWindowListCopyWindowInfo = CGWindowListCopyWindowInfo
                self._kCGWindowListOptionOnScreenOnly = kCGWindowListOptionOnScreenOnly
            except: pass
        if (isinstance(pid, str) and pid.isdigit()) or isinstance(pid, int):
            if main_os == "Windows":
                system_windows = []
                def callback(hwnd, _):
                    if self._win32gui.IsWindowVisible(hwnd):
                        _, window_pid = self._win32process.GetWindowThreadProcessId(hwnd)
                        if window_pid == int(pid): system_windows.append(hwnd)
                self._win32gui.EnumWindows(callback, None)
                return system_windows
            elif main_os == "Darwin":
                system_windows = self._CGWindowListCopyWindowInfo(self._kCGWindowListOptionOnScreenOnly, 0)
                return [win for win in system_windows if win.get("kCGWindowOwnerPID") == int(pid)]
            elif main_os == "Linux":
                try:
                    xdotool_exe = self.getPathFile("/usr/bin/xdotool", "xdotool")
                    if xdotool_exe:
                        res = subprocess.run([xdotool_exe, "search", "--pid", str(pid)], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
                        if res.returncode == 0 and res.stdout.strip(): return [w.strip() for w in res.stdout.strip().splitlines() if w.strip()]
                    wmctrl_exe = self.getPathFile("/usr/bin/wmctrl", "wmctrl")
                    if wmctrl_exe:
                        res = subprocess.run([wmctrl_exe, "-lp"], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
                        if res.returncode == 0: return [line.split()[0] for line in res.stdout.splitlines() if len(line.split()) >= 3 and line.split()[2] == str(pid)]
                except Exception: pass
                return []
            else: return []
        else: return []
    def addToUserPath(self, path_to_add: str):
        if main_os == "Windows":
            if not hasattr(self, "_win32api") or not hasattr(self, "_win32con") or not hasattr(self, "_win32gui"):
                try:
                    try:
                        import win32api # type: ignore
                        import win32con # type: ignore
                        import win32gui # type: ignore
                        self._win32con = win32con
                        self._win32api = win32api
                        self._win32gui = win32gui
                    except Exception:
                        self.install(["pywin32"])
                        self._win32con = self.importModule("win32con")
                        self._win32api = self.importModule("win32api")
                        self._win32gui = self.importModule("win32gui")
                except: pass
            path_to_add = os.path.expandvars(path_to_add)
            current_path = self._win32api.GetEnvironmentVariable("PATH")
            paths = current_path.split(";") if current_path else []
            if path_to_add in paths:
                self.printDebugMessage(f"The path \"{path_to_add}\" was already on the PATH environment variable!")
                return
            paths.append(path_to_add)
            self._win32api.RegSetValueEx(
                self._win32api.RegOpenKeyEx(self._win32con.HKEY_CURRENT_USER, "Environment", 0, self._win32con.KEY_SET_VALUE),
                "PATH",
                0,
                self._win32con.REG_EXPAND_SZ,
                ";".join(paths)
            )
            self._win32gui.SendMessageTimeout(
                self._win32con.HWND_BROADCAST,
                self._win32con.WM_SETTINGCHANGE,
                0,
                "Environment",
                self._win32con.SMTO_ABORTIFHUNG,
                5000
            )
        elif main_os in ("Darwin", "Linux"):
            new_path = os.path.expandvars(os.path.expanduser(path_to_add))
            shell = os.path.basename(os.environ.get("SHELL", "bash" if main_os == "Linux" else "zsh"))
            rc_file = os.path.expanduser(f"~/.{shell}rc")
            export_line = f'export PATH="{new_path}:$PATH"'
            if os.path.exists(rc_file):
                with open(rc_file, "r") as f:
                    if export_line in f.read():
                        self.printDebugMessage(f"The path \"{rc_file}\" was already on the PATH environment variable!")
                        pass
                    else:
                        with open(rc_file, "a") as f_append: f_append.write("\n" + export_line + "\n")
            else:
                with open(rc_file, "w") as f: f.write(export_line + "\n")
            current_path = os.environ.get("PATH", "")
            paths = current_path.split(":")
            if new_path not in paths: os.environ["PATH"] = f"{new_path}:{current_path}"
    def printDebugMessage(self, message: str):
        if self.debug == True: print(f"\033[38;5;226m[PyKits] [DEBUG]: {message}\033[0m")
class plist:
    """
    A class that allows you to save or load plist files.
    """
    __slots__ = ()
    def __init__(self): pass
    def readPListFile(self, path: str):
        if os.path.exists(path):
            with open(path, "rb") as f: plist_data = plistlib.load(f)
            return plist_data
        else: return {}
    def writePListFile(self, path: str, data: typing.Union[dict, str, int, float], binary: bool=False, ns_mode: bool=False):
        try:
            if ns_mode == True and main_os == "Darwin":
                domain = os.path.basename(path).replace(".plist", "", 1)
                for i, v in data.items(): subprocess.run(["defaults", "write", domain, i, str(v)], check=True)
            with open(path, "wb") as f:
                if binary == True: plistlib.dump(data, f, fmt=plistlib.FMT_BINARY)
                else: plistlib.dump(data, f)
            return {"success": True, "message": "Success!", "data": data}
        except Exception as e: return {"success": False, "message": "Something went wrong.", "data": e}
    def dump(self, data: typing.Union[dict, str, int, float], f, binary: bool=False):
        try:
            if binary == True: plistlib.dump(data, f, fmt=plistlib.FMT_BINARY)
            else: plistlib.dump(data, f)
            return {"success": True, "message": "Success!", "data": data}
        except Exception as e: return {"success": False, "message": "Something went wrong.", "data": e}
    def dumps(self, data: typing.Union[dict, str, int, float], binary: bool=False):
        try:
            if binary == True: return plistlib.dumps(data, fmt=plistlib.FMT_BINARY)
            else: return plistlib.dumps(data)
        except Exception as e: return None
    def load(self, f):
        plist_data = plistlib.load(f)
        return plist_data
    loads = readPListFile
class Translator:
    """
    A class that allows you to work with translations from messages to an another message from a translation folder.
    """
    __slots__ = (
        "translation_folder", "start_language", "end_language", "language",
        "delimiter", "translation_json", "patterns", "module_imports",
        "indexed_patterns", "fallback_patterns", "one_group_patterns", "translate"
    )
    DEFAULT_PATTERNS = {
        "- (*)": "- (*)",
        "[(*)] = (*)": "[(*)] = (*)",
        "[(*)] = (*) ((*))": "[(*)] = (*) ((*))",
        "[(*)] [(*)] = (*)": "[(*)] [(*)] = (*)",
        "[*] = (*)": "[*] = (*)",
        "v(*)": "v(*)",
        "(*) | (*) => (*)": "(*) | (*) => (*)",
        "> ": "> ",
        "\u001b[38;5;(*)m(*)\u001b[0m": "\u001b[38;5;(*)m(*)\u001b[0m",
        "\u001b[38;5;(*)m(*)\u001b[0m\n": "\u001b[38;5;(*)m(*)\u001b[0m\n",
        "\u001b[38;5;(*)m(*)\u001b[0m\r\n": "\u001b[38;5;(*)m(*)\u001b[0m\r\n",
    }
    def __init__(self, translation_folder: str, start_language: str=None, end_language: str=None, delimiter: str="\x1f", include_ansi: bool=False):
        self.translation_folder = os.path.abspath(translation_folder)
        self.start_language = start_language or "template"
        self.end_language = end_language if end_language else self.determine_system_language()
        self.language = self.end_language
        self.delimiter = delimiter
        self.translation_json = {}
        self.patterns = {}
        self.module_imports = {}
        self.indexed_patterns = defaultdict(list)
        self.fallback_patterns = []
        self.one_group_patterns = []
        self.translate = functools.lru_cache(maxsize=2048)(self._translate)
        if self.end_language: self.load_new_language(start_lang=self.start_language, end_lang=self.end_language, include_ansi=include_ansi)
    def __bool__(self): return bool(self.end_language) and self.end_language != self.start_language and bool(self.translation_json)
    def _resolve_language_path(self, lang: str) -> str:
        if os.path.isfile(lang): return lang
        cand = os.path.join(self.translation_folder, lang)
        if os.path.isfile(cand): return cand
        if lang == self.start_language:
            template_path = os.path.join(self.translation_folder, "template")
            if os.path.isfile(template_path): return template_path
        raise FileNotFoundError(f"Translation file for '{lang}' not found in '{self.translation_folder}'.")
    def _read_entries(self, path: str) -> list[str]:
        with open(path, "r", encoding="utf-8") as f: content = f.read().replace("\r\n", "\n")
        lines = content.split(self.delimiter)
        return [line.replace("\\n", "\n") for line in lines]
    def load_new_language(self, start_lang: str=None, end_lang: str=None, include_ansi: bool=False):
        if start_lang: self.start_language = start_lang
        self.end_language = end_lang
        self.language = end_lang
        self.translate.cache_clear()
        self.translation_json = {}
        if not end_lang or end_lang == self.start_language or (end_lang == "en" and self.start_language == "template"): return
        start_path = self._resolve_language_path(self.start_language)
        end_path = self._resolve_language_path(self.end_language)
        start_entries = self._read_entries(start_path)
        end_entries = self._read_entries(end_path)
        self.translation_json = { src: dst for src, dst in zip(start_entries, end_entries) if src }
        self.translation_json.update(self.DEFAULT_PATTERNS)
        self.generate_patterns(include_ansi=include_ansi)
    def generate_patterns(self, include_ansi=False):
        self.patterns = {}
        self.indexed_patterns = defaultdict(list)
        self.fallback_patterns = []
        if include_ansi:
            ansi_template = "\033[38;5;(*)m(*)\033[0m\n"
            self.translation_json[ansi_template] = ansi_template
        for template, target in self.translation_json.items():
            if "(*)" not in template: continue
            regex = self.generate_regex(template)
            self.patterns[template] = regex
            parts = template.split("(*)")
            prefix = parts[0]
            suffix = parts[-1]
            split_target = target.split("(*)")
            entry = (template, regex, prefix, suffix, split_target)
            if len(prefix) >= 4: self.indexed_patterns[prefix[:4]].append(entry)
            else: self.fallback_patterns.append(entry)
        self.one_group_patterns = [r for r in self.patterns.values() if r.groups == 1]
    def generate_regex(self, template: str, tolerant: bool=False):
        parts = template.split("(*)")
        escaped = []
        for part in parts:
            if tolerant:
                part = re.sub(r'\.{2,}', r'\\.+', re.escape(part))
                part = re.sub(r'!{2,}', r'!+', part)
            else: part = re.escape(part)
            escaped.append(part)
        pattern = "".join(p + (r"(.+?)" if i < len(escaped) - 1 else "") for i, p in enumerate(escaped))
        return re.compile(f"^{pattern}$")
    def extract_placeholders(self, template: str, actual: str):
        regex = self.patterns.get(template) or self.patterns.setdefault(template, self.generate_regex(template))
        match = regex.match(actual)
        return list(match.groups()) if match else None
    def determine_system_language(self):
        try:
            sys_loc, _ = locale.getlocale()
            if not sys_loc: return "en"
            norm_loc = sys_loc.lower().replace("_", "-")
            base_loc = norm_loc.split("-")[0]
            if os.path.isdir(self.translation_folder):
                available = set(os.listdir(self.translation_folder))
                if norm_loc in available: return norm_loc
                if base_loc in available: return base_loc
        except Exception: pass
        return "en"
    def pre_translate(self, message: str, translate_id: str, match=None, split_target=None):
        try:
            if match is None:
                regex = self.patterns.get(translate_id)
                if not regex: return message
                match = regex.match(message)
                if not match: return message
            tr_placeholders = match.groups()
            sp = split_target if split_target is not None else self.translation_json[translate_id].split("(*)")
            result = []
            for i in range(len(sp)):
                result.append(sp[i])
                if i < len(tr_placeholders):
                    placeholder = tr_placeholders[i]
                    if placeholder in self.translation_json: result.append(self.translation_json[placeholder])
                    elif any(regex.fullmatch(placeholder) for regex in self.one_group_patterns): result.append(self._translate(placeholder))
                    else: result.append(placeholder)
            return "".join(result)
        except Exception: return message
    def _translate(self, message: str): 
        if bool(self):
            exact = self.translation_json.get(message)
            if exact is not None: return exact
            candidates = (self.indexed_patterns.get(message[:4], []) if len(message) >= 4 else []) + self.fallback_patterns
            for template, regex, prefix, suffix, split_target in candidates:
                if prefix and not message.startswith(prefix): continue
                if suffix and not message.endswith(suffix): continue
                match = regex.fullmatch(message)
                if match: return self.pre_translate(message, template, match=match, split_target=split_target)
        return message
class Terminal:
    """
    A master class for replacing Python's stdout and stderr with a more advanced terminal with support for colors, translations (with Translator class) and more!
    """
    ANSI_TO_HEX_TABLE = {
        # 16 bit ANSI
        0: "#000000",  1: "#800000",  2: "#008000",  3: "#808000",
        4: "#000080",  5: "#800080",  6: "#008080",  7: "#c0c0c0",
        8: "#808080",  9: "#ff0000", 10: "#00ff00", 11: "#ffff00",
        12: "#0000ff", 13: "#ff00ff", 14: "#00ffff", 15: "#ffffff",

        # 256 bit ANSI
        16: "#000000", 17: "#00005f", 18: "#000087", 19: "#0000af", 20: "#0000d7", 21: "#0000ff",
        22: "#005f00", 23: "#005f5f", 24: "#005f87", 25: "#005faf", 26: "#005fd7", 27: "#005fff",
        28: "#008700", 29: "#00875f", 30: "#008787", 31: "#0087af", 32: "#0087d7", 33: "#0087ff",
        34: "#00af00", 35: "#00af5f", 36: "#00af87", 37: "#00afaf", 38: "#00afd7", 39: "#00afff",
        40: "#00d700", 41: "#00d75f", 42: "#00d787", 43: "#00d7af", 44: "#00d7d7", 45: "#00d7ff",
        46: "#00ff00", 47: "#00ff5f", 48: "#00ff87", 49: "#00ffaf", 50: "#00ffd7", 51: "#00ffff",
        52: "#5f0000", 53: "#5f005f", 54: "#5f0087", 55: "#5f00af", 56: "#5f00d7", 57: "#5f00ff",
        58: "#5f5f00", 59: "#5f5f5f", 60: "#5f5f87", 61: "#5f5faf", 62: "#5f5fd7", 63: "#5f5fff",
        64: "#5f8700", 65: "#5f875f", 66: "#5f8787", 67: "#5f87af", 68: "#5f87d7", 69: "#5f87ff",
        70: "#5faf00", 71: "#5faf5f", 72: "#5faf87", 73: "#5fafaf", 74: "#5fafd7", 75: "#5fafff",
        76: "#5fd700", 77: "#5fd75f", 78: "#5fd787", 79: "#5fd7af", 80: "#5fd7d7", 81: "#5fd7ff",
        82: "#5fff00", 83: "#5fff5f", 84: "#5fff87", 85: "#5fffaf", 86: "#5fffd7", 87: "#5fffff",
        88: "#870000", 89: "#87005f", 90: "#870087", 91: "#8700af", 92: "#8700d7", 93: "#8700ff",
        94: "#875f00", 95: "#875f5f", 96: "#875f87", 97: "#875faf", 98: "#875fd7", 99: "#875fff",
        100: "#878700", 101: "#87875f", 102: "#878787", 103: "#8787af", 104: "#8787d7", 105: "#8787ff",
        106: "#87af00", 107: "#87af5f", 108: "#87af87", 109: "#87afaf", 110: "#87afd7", 111: "#87afff",
        112: "#87d700", 113: "#87d75f", 114: "#87d787", 115: "#87d7af", 116: "#87d7d7", 117: "#87d7ff",
        118: "#87ff00", 119: "#87ff5f", 120: "#87ff87", 121: "#87ffaf", 122: "#87ffd7", 123: "#87ffff",
        124: "#af0000", 125: "#af005f", 126: "#af0087", 127: "#af00af", 128: "#af00d7", 129: "#af00ff",
        130: "#af5f00", 131: "#af5f5f", 132: "#af5f87", 133: "#af5faf", 134: "#af5fd7", 135: "#af5fff",
        136: "#af8700", 137: "#af875f", 138: "#af8787", 139: "#af87af", 140: "#af87d7", 141: "#af87ff",
        142: "#afaf00", 143: "#afaf5f", 144: "#afaf87", 145: "#afafaf", 146: "#afafd7", 147: "#afafff",
        148: "#afd700", 149: "#afd75f", 150: "#afd787", 151: "#afd7af", 152: "#afd7d7", 153: "#afd7ff",
        154: "#afff00", 155: "#afff5f", 156: "#afff87", 157: "#afffaf", 158: "#afffd7", 159: "#afffff",
        160: "#d70000", 161: "#d7005f", 162: "#d70087", 163: "#d700af", 164: "#d700d7", 165: "#d700ff",
        166: "#d75f00", 167: "#d75f5f", 168: "#d75f87", 169: "#d75faf", 170: "#d75fd7", 171: "#d75fff",
        172: "#d78700", 173: "#d7875f", 174: "#d78787", 175: "#d787af", 176: "#d787d7", 177: "#d787ff",
        178: "#d7af00", 179: "#d7af5f", 180: "#d7af87", 181: "#d7afaf", 182: "#d7afd7", 183: "#d7afff",
        184: "#d7d700", 185: "#d7d75f", 186: "#d7d787", 187: "#d7d7af", 188: "#d7d7d7", 189: "#d7d7ff",
        190: "#d7ff00", 191: "#d7ff5f", 192: "#d7ff87", 193: "#d7ffaf", 194: "#d7ffd7", 195: "#d7ffff",
        196: "#ff0000", 197: "#ff005f", 198: "#ff0087", 199: "#ff00af", 200: "#ff00d7", 201: "#ff00ff",
        202: "#ff5f00", 203: "#ff5f5f", 204: "#ff5f87", 205: "#ff5faf", 206: "#ff5fd7", 207: "#ff5fff",
        208: "#ff8700", 209: "#ff875f", 210: "#ff8787", 211: "#ff87af", 212: "#ff87d7", 213: "#ff87ff",
        214: "#ffaf00", 215: "#ffaf5f", 216: "#ffaf87", 217: "#ffafaf", 218: "#ffafd7", 219: "#ffafff",
        220: "#ffd700", 221: "#ffd75f", 222: "#ffd787", 223: "#ffd7af", 224: "#ffd7d7", 225: "#ffd7ff",
        226: "#ffff00", 227: "#ffff5f", 228: "#ffff87", 229: "#ffffaf", 230: "#ffffd7", 231: "#ffffff",

        # Grayscale
        232: "#080808", 233: "#121212", 234: "#1c1c1c", 235: "#262626",
        236: "#303030", 237: "#3a3a3a", 238: "#444444", 239: "#4e4e4e",
        240: "#585858", 241: "#626262", 242: "#6c6c6c", 243: "#767676",
        244: "#808080", 245: "#8a8a8a", 246: "#949494", 247: "#9e9e9e",
        248: "#a8a8a8", 249: "#b2b2b2", 250: "#bcbcbc", 251: "#c6c6c6",
        252: "#d0d0d0", 253: "#dadada", 254: "#e4e4e4", 255: "#eeeeee"
    }
    SGI_COLOR_TABLE = {
        "Black": [30, 90, 40, 100], "Red": [31, 91, 41, 101], 
        "Green": [32, 92, 42, 102], "Yellow": [33, 93, 43, 103], 
        "Blue": [34, 94, 44, 104], "Magenta": [35, 95, 45, 105], 
        "Teal": [36, 96, 46, 106], "White": [37, 97, 47, 107]
    }

    # Classes
    class CompletedProcess:
        __slots__ = ("stdout", "returncode")
        def __init__(self, stdout, returncode): 
            self.stdout = stdout
            self.returncode = returncode
    class ProgressBarContext:
        "Handles the state of an active progress or timer bar."
        __slots__ = ("terminal", "is_timer", "total", "current", "status_text", "finish_text", "auto_colored", "active")
        def __init__(self, terminal, is_timer=False, total=100, finish_text="Complete", auto_colored=True):
            self.terminal: Terminal = terminal
            self.is_timer = is_timer
            self.total = total
            self.current = total if is_timer else 0
            self.status_text = ""
            self.finish_text = finish_text
            self.auto_colored = auto_colored
            self.active = False
        def start(self):
            self.active = True
            self.terminal._start_progress_mode(self)
            if self.is_timer:
                while self.current > 0 and self.active:
                    self.submit("", self.current)
                    self.current -= 1
                    time.sleep(1)
                self.submit(self.finish_text, 0)
                self.end()
        def submit(self, status: str, value: int):
            self.status_text = status.replace("\033ERR", "❌ ")
            self.current = value
            self.terminal._render_progress()
        def end(self):
            self.active = False
            self.terminal._end_progress_mode()
    class SubprocessStream:
        "A file-like adapter to allow attaching to a subprocess's stdout or stderr and redirecting it to the terminal while still allowing the subprocess to write to it."
        __slots__ = ("terminal", "fd_read", "fd_write", "reader_thread")
        def __init__(self, terminal):
            self.terminal = terminal
            self.fd_read, self.fd_write = os.pipe()
            self.reader_thread = threading.Thread(target=self._reader, daemon=True)
            self.reader_thread.start()
        def _reader(self):
            with os.fdopen(self.fd_read, 'r', encoding='utf-8', errors='replace') as f:
                for line in f: self.terminal.write(line)
        def fileno(self):
            "subprocess calls this to get the real OS file descriptor."
            return self.fd_write
        def close(self):
            "Closing the write end forces the reader thread to hit EOF and gracefully exit."
            try: os.close(self.fd_write)
            except OSError: pass
        def __enter__(self): return self
        def __exit__(self, exc_type, exc_val, exc_tb):
            self.close()
            self.reader_thread.join(timeout=1.0)
    class Color:
        __slots__ = ("rgb", "hex", "ansi", "decimal")
        def __init__(self, r: int, g: int, b: int):
            limit = lambda x: max(0, min(255, int(x)))
            self.rgb = (limit(r), limit(g), limit(b))
            self.hex = Terminal.rgb_to_hex(*self.rgb)
            self.ansi = Terminal.hex_to_ansi(self.hex)
            self.decimal = (self.rgb[0] << 16) + (self.rgb[1] << 8) + self.rgb[2]
        def __int__(self): return self.ansi
        def __str__(self): return f"Color[{self.rgb[0]}, {self.rgb[1]}, {self.rgb[2]}]"
        def wrap(self, msg: str): return Terminal.wrap_rgb(msg, *self.rgb)
        def bg_wrap(self, msg: str): return f"\033[48;2;{self.rgb[0]};{self.rgb[1]};{self.rgb[2]}m{msg}{Terminal.get_reset_color()}"
        def grayscale(self): 
            gray = int(0.299 * self.rgb[0] + 0.587 * self.rgb[1] + 0.114 * self.rgb[2])
            return Terminal.Color(gray, gray, gray)
    class Black(Color):
        def __init__(self): super().__init__(0, 0, 0)
    class Red(Color):
        def __init__(self): super().__init__(255, 0, 0)
    class Green(Color):
        def __init__(self): super().__init__(0, 255, 0)
    class Yellow(Color):
        def __init__(self): super().__init__(255, 255, 0)
    class Blue(Color):
        def __init__(self): super().__init__(0, 0, 255)
    class Magenta(Color):
        def __init__(self): super().__init__(255, 0, 255)
    class Teal(Color):
        def __init__(self): super().__init__(0, 255, 255)
    class White(Color):
        def __init__(self): super().__init__(255, 255, 255)

    # Initialization
    def __init__(self, logger: logging.Logger=None, log_level: int=None, target_language: str=None, tail_lines: int=3, translator_path: str=None):
        self.logger = logger
        self.log_level = log_level
        self.encoding = "utf-8"
        self.buffer = ""
        self.line_count = 0
        
        # Translation
        if not translator_path: translator_path = os.path.join(cur_path, "Translations")
        self.translator = Translator(translator_path, start_language=None, end_language=target_language, delimiter="\x1f")
        self.translate = self.translator.translate
        
        # Progress Settings
        self.tail_lines = tail_lines
        self.progress_active = False
        self.active_bar = None
        self.queued_logs = []
        self.tail_buffer = collections.deque(maxlen=tail_lines)
        self._fix_windows_ansi()

        # Static Colors
        for clr in (self.Black, self.Red, self.Green, self.Yellow, self.Blue, self.Magenta, self.Teal, self.White): 
            setattr(self, clr.__name__.upper(), clr())
    def _fix_windows_ansi(self):
        if os.name == "nt":
            try:
                kernel32 = ctypes.windll.kernel32
                handle = kernel32.GetStdHandle(-11)
                mode = ctypes.c_uint()
                kernel32.GetConsoleMode(handle, ctypes.byref(mode))
                kernel32.SetConsoleMode(handle, mode.value | 0x0004)
            except Exception: pass
    def inject_stdout(self):
        "Replaces sys.stdout with this Terminal instance."
        sys.stdout = self
    def inject_stderr(self):
        "Replaces sys.stderr with this Terminal instance."
        sys.stderr = self
    def inject_full_terminal(self):
        "Replaces the stdout, stderr and input functions in python"
        sys.stdout = self
        self._original_input = builtins.input
        builtins.input = self.input
    def restore_stdout(self):
        "Restores sys.stdout to the original stdout."
        sys.stdout = sys.__stdout__
    def restore_stderr(self):
        "Restores sys.stderr to the original stderr."
        sys.stderr = sys.__stderr__
    def restore_full_terminal(self):
        "Restores the stdout, stderr and input functions in python (if stored in this class/object)"
        sys.stdout = sys.__stdout__
        sys.stderr = sys.__stderr__
        if hasattr(self, "_original_input"): builtins.input = self._original_input

    # Stdout Replacement Methods
    @staticmethod
    def _raw_stdout_write(text: str):
        if sys.__stdout__ is not None:
            try: sys.__stdout__.write(text)
            except Exception: pass
    @staticmethod
    def _raw_stdout_flush():
        if sys.__stdout__ is not None:
            try: sys.__stdout__.flush()
            except Exception: pass
    def fileno(self):
        if sys.__stdout__ is not None: return sys.__stdout__.fileno()
        raise OSError("No underlying stdout fileno in GUI/service mode")
    def write(self, message: str):
        message = self.translate(message)
        self.buffer += message
        if "\n" in self.buffer:
            lines = self.buffer.splitlines(keepends=True)
            if not lines[-1].endswith("\n"): self.buffer = lines.pop()
            else: self.buffer = ""
            if self.progress_active and lines:
                prev_lines = getattr(self, "_prev_bar_lines", 0)
                if prev_lines > 0:
                    if prev_lines > 1: self._raw_stdout_write(f"\r\033[{prev_lines - 1}A\033[J")
                    else: self._raw_stdout_write("\r\033[J")
                self._prev_bar_lines = 0
            for line in lines:
                self.line_count += 1
                clean_line = line.rstrip()
                if clean_line and self.logger and self.log_level is not None:
                    try: self.logger.log(self.log_level, clean_line)
                    except Exception: self.logger.log(self.log_level, clean_line.encode(self.encoding, errors="replace").decode(self.encoding))
                self._raw_stdout_write(line)
            if self.progress_active and lines: self._render_progress()
        self._raw_stdout_flush()
    def flush(self):
        if self.buffer:
            if self.progress_active:
                prev_lines = getattr(self, "_prev_bar_lines", 0)
                if prev_lines > 0:
                    if prev_lines > 1: self._raw_stdout_write(f"\r\033[{prev_lines - 1}A\033[J")
                    else: self._raw_stdout_write("\r\033[J")
                self._prev_bar_lines = 0
            if self.buffer.rstrip() and self.logger and self.log_level is not None:
                try: self.logger.log(self.log_level, self.buffer.rstrip())
                except Exception: pass
            self._raw_stdout_write(self.buffer)
            if self.progress_active: self._render_progress()
            else: self._raw_stdout_flush()
        self.buffer = ""
    def clear(self):
        "Clears the console completely."
        self._raw_stdout_write("\033[2J\033[3J\033[H" if os.name == "nt" else "\033c\033[3J")
        self._raw_stdout_flush()
        self.line_count = 0
    def input(self, prompt=""):
        "Requests an input from the user using prompt, similar to input()"
        translated_prompt = self.translate(str(prompt))
        sys.__stdout__.write(translated_prompt)
        sys.__stdout__.flush()
        return sys.stdin.readline().rstrip('\n')
    
    # Printing Methods
    def print_main(self, message: str):
        "Standard white console message."
        return self.print_ansi(message, 255)
    def print_success(self, message: str):
        "Standard green success console message."
        return self.print_ansi(message, 82)
    def print_error(self, message: str):
        "Standard red fail console message."
        return self.print_ansi(message, 196)
    def print_warning(self, message: str):
        "Standard yellow warning console message."
        return self.print_ansi(message, 226)
    def print_debug(self, message: str):
        "Standard grey debug console message."
        return self.print_ansi(message, 245)
    def print_system(self, message: str):
        "Standard blue title console message."
        return self.print_ansi(message, 51)
    def print_ansi(self, message: str, ansi: int): print(self.wrap_ansi(message, ansi))
    def print_rgb(self, message: str, r: int, g: int, b: int): print(self.wrap_rgb(message, r, g, b))
    def print_color(self, message: str, color: Color): print(self.wrap_color(message, color))
    def print_gradient(self, message: str, color_stops: typing.List[str], angle: float = 0.0): print(self.wrap_gradient(message, color_stops, angle))
    def change_last_line(self, message: str):
        message = self.translate(message)
        if self.progress_active:
            prev_lines = getattr(self, '_prev_bar_lines', 0)
            if prev_lines > 0:
                if prev_lines > 1: sys.__stdout__.write(f"\r\033[{prev_lines - 1}A\033[J")
                else: sys.__stdout__.write("\r\033[J")
            self._prev_bar_lines = 0
        sys.__stdout__.write(f"\033[1A\033[2K\r{message}\n")
        if message.strip() and self.logger and self.log_level is not None:
            try: self.logger.log(self.log_level, message.strip())
            except Exception: pass
        if self.progress_active: self._render_progress()
        else: sys.__stdout__.flush()
    def change_last_lines(self, messages: typing.List[str]):
        num_lines = len(messages)
        if num_lines == 0: return
        translated = [self.translate(m) for m in messages]
        if self.progress_active:
            prev_lines = getattr(self, '_prev_bar_lines', 0)
            if prev_lines > 0:
                if prev_lines > 1: sys.__stdout__.write(f"\r\033[{prev_lines - 1}A\033[J")
                else: sys.__stdout__.write("\r\033[J")
            self._prev_bar_lines = 0
        sys.__stdout__.write(f"\033[{num_lines}A")
        for msg in translated:
            sys.__stdout__.write(f"\033[2K\r{msg}\n")
            if msg.strip() and self.logger and self.log_level is not None:
                try: self.logger.log(self.log_level, msg.strip())
                except Exception: pass
        if self.progress_active: self._render_progress()
        else: sys.__stdout__.flush()
    def clear_last_lines(self, num_lines: int):
        if num_lines <= 0: return
        if self.progress_active:
            prev_lines = getattr(self, '_prev_bar_lines', 0)
            if prev_lines > 0:
                if prev_lines > 1: sys.__stdout__.write(f"\r\033[{prev_lines - 1}A\033[J")
                else: sys.__stdout__.write("\r\033[J")
            self._prev_bar_lines = 0
        sys.__stdout__.write(f"\033[{num_lines}A")
        for _ in range(num_lines): sys.__stdout__.write("\033[2K\r\n")
        sys.__stdout__.write(f"\033[{num_lines}A")
        if self.progress_active: self._render_progress()
        else: sys.__stdout__.flush()

    # Translation
    def change_language(self, start_language: str=None, end_language: str=None, include_ansi: bool=False): self.translator.load_new_language(start_lang=start_language, end_lang=end_language, include_ansi=include_ansi)
    def change_translator(self, translator: Translator): 
        self.translator = translator
        self.translate = self.translator.translate
    def stdout_translate(self, message: str):
        mes = str(message)
        if hasattr(sys.stdout, "translate"): mes = sys.stdout.translate(mes)
        return mes

    # Progress Bars
    def create_progress_bar(self) -> ProgressBarContext: return self.ProgressBarContext(self, is_timer=False, total=100)
    def create_timer_bar(self, seconds: int, finish_text: str="Done!") -> ProgressBarContext: return self.ProgressBarContext(self, is_timer=True, total=seconds, finish_text=finish_text)
    def _start_progress_mode(self, bar: ProgressBarContext):
        self.progress_active = True
        self.active_bar = bar
        self.queued_logs.clear()
        self._prev_bar_lines = 0
    def _render_progress(self):
        if not self.active_bar: return
        bar = self.active_bar
        now = time.time()
        if not hasattr(self, "_last_term_width_time") or now - getattr(self, "_last_term_width_time") > 1:
            self._last_term_width = shutil.get_terminal_size((80, 20)).columns
            self._last_term_width_time = now
        term_width = getattr(self, "_last_term_width")
        ansi_escape = getattr(self.__class__, "_ansi_escape", None)
        if not ansi_escape:
            ansi_escape = re.compile(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])')
            self.__class__._ansi_escape = ansi_escape
        output_buffer = []
        prev_lines = getattr(self, '_prev_bar_lines', 0)
        if prev_lines > 0:
            if prev_lines > 1: output_buffer.append(f"\r\033[{prev_lines - 1}A\033[J")
            else: output_buffer.append("\r\033[J")
        self._prev_bar_lines = 0
        is_error = "❌" in bar.status_text or "\033ERR" in bar.status_text
        display_status = bar.status_text.replace("\033ERR", "").replace("❌ ", "").strip()
        clean_status = ansi_escape.sub('', display_status)
        is_done = (bar.current <= 0) if bar.is_timer else (bar.current >= 100)
        theme_color = 196 if is_error else (82 if is_done else 255)
        color_code = f"\033[38;5;{theme_color}m" if bar.auto_colored else ""
        reset_code = "\033[0m" if bar.auto_colored else ""
        if bar.auto_colored: display_status = display_status.replace("\033[0m", f"\033[0m{color_code}")
        if bar.is_timer:
            fin = round(((bar.current / bar.total) * 100) / (100 / bar.total)) if bar.total > 0 else 100
            prefix = "✅" if is_done else "⏰"
            filled_blocks = '█' * int(max(0, fin))
            empty_blocks = ' ' * int(max(0, bar.total - fin))
            bar_visual = f"[{filled_blocks}{empty_blocks}]"
            status_or_time = f"{bar.current}s"
            combined_clean = f"{prefix} {bar_visual} {status_or_time}"
            if len(combined_clean) > term_width - 2:
                bar_lines_text = [
                    f"{color_code}{prefix} {status_or_time}{reset_code}",
                    f"{color_code}{bar_visual}{reset_code}"
                ]
            else: bar_lines_text = [f"{color_code}{prefix} {bar_visual} {status_or_time}{reset_code}"]
            if is_done and display_status: bar_lines_text.append(f"\033[0m{display_status}")
        else:
            fin = round(bar.current * 0.2)
            prefix = "✅" if is_done else ("❌" if is_error else "🚀")
            filled_blocks = '█' * int(max(0, fin))
            empty_blocks = ' ' * int(max(0, 20 - fin))
            bar_visual = f"[{filled_blocks}{empty_blocks}]"
            percent_str = f"{int(bar.current)}%"
            combined_clean = f"{prefix} {clean_status} {bar_visual} {percent_str}"
            if len(combined_clean) > term_width - 2:
                bar_lines_text = [
                    f"{color_code}{prefix} {display_status} {percent_str}{reset_code}",
                    f"{color_code}{bar_visual}{reset_code}"
                ]
            else: bar_lines_text = [f"{color_code}{prefix} {display_status} {bar_visual} {percent_str}{reset_code}"]
        def calc_lines(t):
            clean = ansi_escape.sub('', t)
            if hasattr(sys, "_app_delegate"): return clean.count('\n') + 1
            total = 0
            for line in clean.split('\n'): total += 1 if len(line) == 0 else max(1, math.ceil(len(line) / term_width))
            return total
        current_bar_lines = sum(calc_lines(t) for t in bar_lines_text)
        for i, text in enumerate(bar_lines_text):
            if i < len(bar_lines_text) - 1: output_buffer.append(f"{text}\033[K\n")
            else: output_buffer.append(f"{text}\033[K")
        self._prev_bar_lines = current_bar_lines
        sys.__stdout__.write("".join(output_buffer))
        sys.__stdout__.flush()
    def _end_progress_mode(self):
        self.progress_active = False
        self.active_bar = None
        sys.__stdout__.write("\n")
        sys.__stdout__.flush()

    # Running Processes Under Console
    def get_stream(self) -> SubprocessStream: return self.SubprocessStream(self)
    def run_process(self, args=["python3"], cwd=None):
        output = []
        if os.name != "nt":
            pid, fd = pty.fork()
            if pid == 0: 
                if cwd: os.chdir(cwd)
                os.execvp(args[0], args)
            else:
                try:
                    while True:
                        rlist, _, _ = select.select([fd, sys.stdin], [], [])
                        if fd in rlist:
                            try:
                                data = os.read(fd, 1024)
                                if not data: break
                                output.append(data)
                                self.write(data.decode(errors="replace"))
                            except OSError: break
                        if sys.stdin in rlist:
                            user_input = os.read(sys.stdin.fileno(), 1024)
                            if not user_input: break
                            os.write(fd, user_input)
                finally:
                    os.close(fd)
                    try: _, status = os.waitpid(pid, 0)
                    except ChildProcessError: status = 0
                return self.CompletedProcess(b''.join(output).decode().strip(), os.WEXITSTATUS(status))
        else:
            proc = subprocess.Popen(
                args, stdin=subprocess.PIPE, stdout=subprocess.PIPE, 
                stderr=subprocess.STDOUT, cwd=cwd, bufsize=1, universal_newlines=True
            )
            def reader():
                for line in proc.stdout:
                    output.append(line.encode())
                    self.write(line)
            thread = threading.Thread(target=reader)
            thread.start()
            try:
                while proc.poll() is None:
                    user_input = sys.stdin.readline()
                    if user_input: proc.stdin.write(user_input); proc.stdin.flush()
            except KeyboardInterrupt: proc.terminate()
            thread.join()
            return self.CompletedProcess(b''.join(output).decode().strip(), proc.wait())

    # Console Formatting
    @staticmethod
    def cursor_up(lines: int=1): return f"\033[{lines}A"
    @staticmethod
    def cursor_down(lines: int=1): return f"\033[{lines}B"
    @staticmethod
    def cursor_right(cols: int=1): return f"\033[{cols}C"
    @staticmethod
    def cursor_left(cols: int=1): return f"\033[{cols}D"
    @staticmethod
    def cursor_to_start(): return "\r"
    @staticmethod
    def clear_line(): return "\033[2K\r"
    @staticmethod
    def save_cursor(): return "\033[s"
    @staticmethod
    def restore_cursor(): return "\033[u"
    @staticmethod
    def hide_cursor(): return "\033[?25l"
    @staticmethod
    def show_cursor(): return "\033[?25h"
    @staticmethod
    def set_console_title(title: str):
        sys.__stdout__.write(f"\033]0;{title}\007")
        sys.__stdout__.flush()
    @staticmethod
    def get_reset_color(): return "\033[0m"
    @staticmethod
    def bold(msg: str): return f"\033[1m{msg}\033[0m"
    @staticmethod
    def italic(msg: str): return f"\033[3m{msg}\033[0m"
    @staticmethod
    def underline(msg: str): return f"\033[4m{msg}\033[0m"
    @staticmethod
    def strikethrough(msg: str): return f"\033[9m{msg}\033[0m"
    @classmethod
    def foreground(cls, msg: str, color: str="White", bright: bool=False): 
        sgr = cls.SGI_COLOR_TABLE.get(color, cls.SGI_COLOR_TABLE["White"])
        return f"\033[{sgr[1 if bright else 0]}m{msg}{cls.get_reset_color()}"
    @classmethod
    def background(cls, msg: str, color: str="White", bright: bool=False): 
        sgr = cls.SGI_COLOR_TABLE.get(color, cls.SGI_COLOR_TABLE["White"])
        return f"\033[{sgr[3 if bright else 2]}m{msg}{cls.get_reset_color()}"
    @classmethod
    def wrap_ansi(cls, msg: str, ansi_num: int): return f"\033[38;5;{ansi_num}m{msg}{cls.get_reset_color()}"
    @classmethod
    def wrap_rgb(cls, msg: str, r: int, g: int, b: int): return f"\033[38;2;{r};{g};{b}m{msg}{cls.get_reset_color()}"
    @classmethod
    def wrap_color(cls, msg: str, color: Color): return f"\033[38;2;{color.rgb[0]};{color.rgb[1]};{color.rgb[2]}m{msg}{cls.get_reset_color()}"

    # Color Math
    @staticmethod
    def hex_to_rgb(hex_code: str): 
        h = hex_code.lstrip("#")
        return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))
    @classmethod
    def rgb_to_hex(cls, r: int, g: int, b: int): return f"#{max(0, min(255, r)):02x}{max(0, min(255, g)):02x}{max(0, min(255, b)):02x}"
    @classmethod
    def hex_to_color(cls, hex_code: str): 
        r, g, b = cls.hex_to_rgb(hex_code)
        return cls.Color(r, g, b)
    @classmethod
    def hex_to_ansi(cls, hex_code: str):
        r, g, b = cls.hex_to_rgb(hex_code)
        if not hasattr(cls, "_ansi_to_rgb"):
            cls._ansi_to_rgb = {c: cls.hex_to_rgb(h) for c, h in cls.ANSI_TO_HEX_TABLE.items()}
            cls._hex_to_ansi_cache = {}
        if hex_code in cls._hex_to_ansi_cache: return cls._hex_to_ansi_cache[hex_code]
        closest_code, closest_dist = None, float("inf")
        for c, (cr, cg, cb) in cls._ansi_to_rgb.items():
            dist = (cr - r)**2 + (cg - g)**2 + (cb - b)**2
            if dist < closest_dist: 
                closest_code, closest_dist = c, dist
                if dist == 0: break     
        cls._hex_to_ansi_cache[hex_code] = closest_code
        return closest_code
    @staticmethod
    def gamma_correct(value: int): return (value / 255.0) ** 2.2
    @staticmethod
    def inverse_gamma(value: float): return int((value ** (1.0 / 2.2)) * 255.0)
    @classmethod
    @functools.lru_cache(maxsize=128)
    def _generate_gradient_lut(cls, color_stops_tuple):
        lut_size = 256
        gradient_lut = []
        num_segments = len(color_stops_tuple) - 1
        stops_gamma = [(cls.gamma_correct(r), cls.gamma_correct(g), cls.gamma_correct(b)) for r, g, b in [cls.hex_to_rgb(c) for c in color_stops_tuple]]
        for i in range(lut_size):
            t = i / (lut_size - 1)
            scaled_t = t * num_segments
            seg_index = min(int(scaled_t), num_segments - 1)
            local_t = scaled_t - seg_index
            start_rgb, end_rgb = stops_gamma[seg_index], stops_gamma[seg_index + 1]
            r = cls.inverse_gamma(start_rgb[0] + (end_rgb[0] - start_rgb[0]) * local_t)
            g = cls.inverse_gamma(start_rgb[1] + (end_rgb[1] - start_rgb[1]) * local_t)
            b = cls.inverse_gamma(start_rgb[2] + (end_rgb[2] - start_rgb[2]) * local_t)
            gradient_lut.append(f"\033[38;2;{max(0, min(255, r))};{max(0, min(255, g))};{max(0, min(255, b))}m")
        return gradient_lut
    @classmethod
    def wrap_gradient(cls, message: str, color_stops: typing.List[str], angle: float = 0.0):
        if not message or not color_stops: return message
        if len(color_stops) == 1: return cls.wrap_rgb(message, *cls.hex_to_rgb(color_stops[0]))
        gradient_lut = cls._generate_gradient_lut(tuple(color_stops))
        lut_size = len(gradient_lut)
        lines = message.split("\n")
        rad = math.radians(angle)
        dx, dy = math.sin(rad), -math.cos(rad)
        aspect_ratio = 2.0 
        min_proj, max_proj = float('inf'), float('-inf')
        for y, line in enumerate(lines):
            y_proj = (y * aspect_ratio) * dy
            for x, char in enumerate(line):
                if char.strip():
                    proj = x * dx + y_proj
                    if proj < min_proj: min_proj = proj
                    if proj > max_proj: max_proj = proj
        if max_proj == min_proj: return cls.wrap_rgb(message, *cls.hex_to_rgb(color_stops[0]))
        result = []
        for y, line in enumerate(lines):
            line_result = []
            y_proj = (y * aspect_ratio) * dy
            last_lut_ind = None
            for x, char in enumerate(line):
                if not char.strip():
                    line_result.append(char)
                    continue
                proj = x * dx + y_proj
                t = max(0.0, min(1.0, (proj - min_proj) / (max_proj - min_proj)))
                lut_index = int(t * (lut_size - 1))
                if lut_index != last_lut_ind:
                    line_result.append(f"{gradient_lut[lut_index]}{char}")
                    last_lut_ind = lut_index
                else:
                    line_result.append(char)
            result.append("".join(line_result) + cls.get_reset_color())
        return "\n".join(result)
class FileSelector:
    __slots__ = ("unable_to_use_tkinter", "_tk", "_filedialog")
    class Response:
        __slots__ = ("ok", "path")
        def __init__(self, ok: bool, path: typing.Optional[str]):
            self.ok = ok
            self.path = path
    def __init__(self):
        self.unable_to_use_tkinter = False
        pip_class = pip()
        try:
            tk = pip_class.importModule("tkinter")
            filedialog = pip_class.importModule("tkinter.filedialog")
        except Exception as e:
            print(e)
            try:
                pip_class.install(["tk"])
                tk = pip_class.importModule("tkinter")
                filedialog = pip_class.importModule("tkinter.filedialog")
            except Exception as e: self.unable_to_use_tkinter = True; print(e)
        if not self.unable_to_use_tkinter:
            self._tk = tk
            self._filedialog = filedialog
    def _run_in_subprocess(self, python_code: str, is_multiple: bool=False):
        full_code = f"import json; import tkinter as tk; from tkinter import filedialog; root = tk.Tk(); root.withdraw(); {python_code}; print(json.dumps(path))"
        s = subprocess.run([sys.executable, "-c", full_code], capture_output=True, text=True)
        if s.returncode == 0 and s.stdout.strip(): 
            try: result = json.loads(s.stdout.strip())
            except json.JSONDecodeError: return [self.Response(ok=False, path=None)] if is_multiple else self.Response(ok=False, path=None)
            if is_multiple:
                if not result: return [self.Response(ok=False, path=None)]
                return [self.Response(ok=True, path=i) for i in result]
            elif result and isinstance(result, str) and os.path.exists(result): return self.Response(ok=True, path=result)
        return [self.Response(ok=False, path=None)] if is_multiple else self.Response(ok=False, path=None)
    def _get_darwin_b64_args(self, args_dict: dict) -> str:
        b64_encoded = base64.b64encode(json.dumps(args_dict).encode('utf-8')).decode('utf-8')
        return f"import base64; inf=json.loads(base64.b64decode('{b64_encoded}').decode('utf-8'))"
    def select_file(self, title: str="Select a file", initialdir: typing.Optional[str]=None, filetypes: typing.Tuple=(("All files", "*.*"),)):
        if self.unable_to_use_tkinter: return self.Response(ok=False, path=None)
        if main_os == "Darwin": 
            setup_code = self._get_darwin_b64_args({"title": title, "initialdir": initialdir, "filetypes": list(filetypes)})
            return self._run_in_subprocess(f"{setup_code}; path = filedialog.askopenfilename(**inf)")
        root = self._tk.Tk()
        root.withdraw()
        path = self._filedialog.askopenfilename(title=title, initialdir=initialdir, filetypes=filetypes)
        root.destroy()
        if path: return self.Response(ok=True, path=path)
        return self.Response(ok=False, path=None)
    def select_files(self, title: str="Select a file", initialdir: typing.Optional[str]=None, filetypes: typing.Tuple=(("All files", "*.*"),)):
        if self.unable_to_use_tkinter: return False, [self.Response(ok=False, path=None)]
        if main_os == "Darwin": 
            setup_code = self._get_darwin_b64_args({"title": title, "initialdir": initialdir, "filetypes": list(filetypes)})
            s = self._run_in_subprocess(f"{setup_code}; path = filedialog.askopenfilenames(**inf)", is_multiple=True)
            if s and s[0].ok: return True, s
            return False, s
        root = self._tk.Tk()
        root.withdraw()
        path = self._filedialog.askopenfilenames(title=title, initialdir=initialdir, filetypes=filetypes)
        root.destroy()
        if path: return True, [self.Response(ok=True, path=p) for p in path]
        return False, [self.Response(ok=False, path=None)]
    def select_folder(self, title: str="Select a folder", initialdir: typing.Optional[str]=None, mustexist: bool=True):
        if self.unable_to_use_tkinter: return self.Response(ok=False, path=None)
        if main_os == "Darwin": 
            setup_code = self._get_darwin_b64_args({"title": title, "initialdir": initialdir, "mustexist": mustexist})
            return self._run_in_subprocess(f"{setup_code}; path = filedialog.askdirectory(**inf)")
        root = self._tk.Tk()
        root.withdraw()
        path = self._filedialog.askdirectory(title=title, initialdir=initialdir, mustexist=mustexist)
        root.destroy()
        if path: return self.Response(ok=True, path=path)
        return self.Response(ok=False, path=None)
class Socket:
    """
    A class that provides a simple interface for working with data between apps.
    """
    __slots__ = (
        "host", "port", "debug", "topics", "_listener_thread", 
        "_buffer_size", "_running", "_executor"
    )
    def __init__(self, host="127.0.0.1", port=60153):
        self.host = host
        self.port = port
        self.debug = False
        self.topics = {}
        self._listener_thread = None
        self._buffer_size = 65536
        self._running = False
        self._executor = None
    def subscribe(self, topic_name: str, call_func: typing.Callable): self.topics[topic_name] = call_func
    def listen(self):
        if self._running: return
        self._running = True
        if self._executor is None: self._executor = concurrent.futures.ThreadPoolExecutor(max_workers=25)
        self._listener_thread = threading.Thread(target=self._listen_loop, daemon=True)
        self._listener_thread.start()
        self._print_debug(f"System listening on port {self.port}...")
    def close(self):
        self._running = False
        if self._listener_thread and self._listener_thread.is_alive(): self._listener_thread.join()
        if self._executor is not None:
            self._executor.shutdown(wait=False)
            self._executor = None
        self._print_debug("Socket closed as per request.")
    def send(self, topic: str, data: typing.Any):
        payload_dict = {
            "topic": topic,
            "data": data
        }
        payload_bytes = json.dumps(payload_dict).encode('utf-8')
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.setsockopt(socket.IPPROTO_TCP, socket.TCP_NODELAY, 1)
                s.connect((self.host, self.port))
                s.sendall(payload_bytes)
        except ConnectionRefusedError: self._print_debug(f"Could not send \"{topic}\" notification. Is the server running?")
    def request(self, topic: str, data: typing.Any, timeout: float = 5.0):
        payload_dict = {
            "topic": topic,
            "data": data
        }
        payload_bytes = json.dumps(payload_dict).encode('utf-8')
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.setsockopt(socket.IPPROTO_TCP, socket.TCP_NODELAY, 1)
                if timeout: s.settimeout(timeout)
                s.connect((self.host, self.port))
                s.sendall(payload_bytes)
                s.shutdown(socket.SHUT_WR)
                chunks = []
                while True:
                    chunk = s.recv(self._buffer_size)
                    if not chunk: break
                    chunks.append(chunk)
                raw_bytes = b"".join(chunks)
                if raw_bytes:
                    response_payload = json.loads(raw_bytes.decode("utf-8"))
                    return response_payload.get("data")
        except ConnectionRefusedError: self._print_debug(f"Could not request \"{topic}\". Is the server running?")
        except Exception as e: self._print_debug(f"Request error: {e}")
        return None
    def exists(self):
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as test: test.bind((self.host, self.port))
            return False
        except OSError as e:
            if e.errno in (errno.EADDRINUSE, errno.EACCES, 10048, 10013): return True
            raise e
    def wait_till_free(self, timeout: float=None, interval: float=1.0):
        start_time = time.time()
        while self.exists():
            self._print_debug(f"Port {self.port} is in use. Waiting for it to close...")
            if timeout is not None and (time.time() - start_time) > timeout:
                self._print_debug(f"Timeout of {timeout}s reached waiting for port {self.port}.")
                return False
            time.sleep(interval)
        self._print_debug(f"Port {self.port} is available.")
        return True
    def _listen_loop(self):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if hasattr(socket, "SO_EXCLUSIVEADDRUSE"): s.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
            else: s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            s.setsockopt(socket.IPPROTO_TCP, socket.TCP_NODELAY, 1)
            s.bind((self.host, self.port))
            s.listen()
            while self._running:
                rlist, _, _ = select.select([s], [], [], 1.0)
                if not rlist: continue 
                try:
                    conn, addr = s.accept()
                    conn.setsockopt(socket.IPPROTO_TCP, socket.TCP_NODELAY, 1)
                    self._executor.submit(self._handle_responding, conn)
                except Exception as e: self._print_debug(f"Listener error: {e}")
    def _handle_responding(self, conn):
        with conn:
            chunks = []
            while True:
                try:
                    chunk = conn.recv(self._buffer_size)
                    if not chunk: break
                    chunks.append(chunk)
                except Exception as e: self._print_debug(f"Error receiving data: {e}"); break
            raw_bytes = b"".join(chunks)
            if raw_bytes: 
                response_data = self._proc_message(raw_bytes)
                response_payload = {"topic": "_response", "data": response_data}
                try: conn.sendall(json.dumps(response_payload).encode('utf-8'))
                except Exception as e: self._print_debug(f"Could not send response back: {e}")
    def _proc_message(self, data: bytes):
        try:
            payload = json.loads(data.decode('utf-8'))
            topic = payload.get("topic")
            actual_data = payload.get("data")
            if topic in self.topics: return self.topics[topic](actual_data)
            else: self._print_debug(f"Warning: Received message for unknown topic \"{topic}\".")
        except json.JSONDecodeError: self._print_debug("Received invalid JSON data.")
    def _print_debug(self, message):
        if self.debug: print(f"Debug: {message}")
class Lock:
    """
    A class that provides a way to create locks that clear when OS is restarted.
    """
    __slots__ = ("file_name", "file_handle", "_is_windows")
    def __init__(self, file_name: str):
        self.file_name = file_name
        self.file_handle = None
        self._is_windows = os.name == "nt"
    def acquire(self, timeout: float=None) -> bool:
        self.file_handle = open(self.file_name, "w")
        start_time = time.time()
        while True:
            try:
                if self._is_windows: msvcrt.locking(self.file_handle.fileno(), msvcrt.LK_NBLCK, 1)
                else: fcntl.flock(self.file_handle, fcntl.LOCK_EX | fcntl.LOCK_NB)
                return True
            except (IOError, OSError):
                if timeout is not None and (time.time() - start_time) >= timeout:
                    self.file_handle.close()
                    self.file_handle = None
                    return False
                time.sleep(0.1)
    def release(self):
        if self.file_handle is not None:
            try:
                if self._is_windows:
                    self.file_handle.seek(0)
                    msvcrt.locking(self.file_handle.fileno(), msvcrt.LK_UNLCK, 1)
                else: fcntl.flock(self.file_handle, fcntl.LOCK_UN)
            finally:
                self.file_handle.close()
                self.file_handle = None
                try: os.remove(self.file_name)
                except OSError: pass
    def exists(self):
        if not os.path.exists(self.file_name): return False
        try:
            with open(self.file_name, "a") as test_file:
                if self._is_windows:
                    msvcrt.locking(test_file.fileno(), msvcrt.LK_NBLCK, 1)
                    test_file.seek(0)
                    msvcrt.locking(test_file.fileno(), msvcrt.LK_UNLCK, 1)
                else:
                    fcntl.flock(test_file, fcntl.LOCK_EX | fcntl.LOCK_NB)
                    fcntl.flock(test_file, fcntl.LOCK_UN)
            return False
        except (IOError, OSError): return True
    def wait_till_free(self, timeout: float=None, check_interval: float=0.1):
        start_time = time.time()
        while self.exists():
            if timeout is not None and (time.time() - start_time) >= timeout: return False
            time.sleep(check_interval)
        return True
    def __enter__(self):
        if not self.acquire(timeout=10): raise TimeoutError(f"Could not acquire lock on {self.file_name}")
        return self
    def __exit__(self, exc_type, exc_val, exc_tb): self.release()
class ProgressBar:
    """
    A class that allows you to work with progress bars in the console using the Terminal class.
    """
    __slots__ = ("terminal", "context", "current_percentage", "status_text")
    def __init__(self, terminal_instance: Terminal=None):
        self.terminal = terminal_instance if terminal_instance else Terminal()
        self.context = self.terminal.create_progress_bar()
        self.current_percentage = 0
        self.status_text = ""
    def __int__(self): return int(self.current_percentage)
    def submit(self, status_text: str, percentage: int):
        self.current_percentage = percentage
        self.status_text = status_text
        self.context.submit(status_text, percentage)
    def start(self): self.context.start()
    def end(self): self.context.end()
    def error(self): return "\033ERR"
class TimerBar(ProgressBar):
    """
    A class that allows you to work with countdown timers in the console.
    """
    __slots__ = ("current_countdown", "started", "finished_text", "begin_in_end")
    def __init__(self, terminal_instance: Terminal=None, countdown: int=5, finished_text: str="Continue with your action!", begin_in_end: bool=True):
        self.terminal = terminal_instance if terminal_instance else Terminal()
        self.context = self.terminal.create_timer_bar(seconds=countdown, finish_text=finished_text)
        self.current_countdown = int(countdown)
        self.started = int(countdown)
        self.finished_text = finished_text
        self.begin_in_end = begin_in_end   
    def __int__(self): return int(self.current_countdown)
    def submit(self):
        status = self.finished_text if self.current_countdown == 0 else f"{self.current_countdown}s remaining"
        self.context.submit(status, self.current_countdown)
    def start(self): self.context.start()
class InstantRequestJSONResponse:
    """
    A class that represents an instant JSON response.
    """
    __slots__ = ("json", "ok")
    def __init__(self, data: typing.Dict=None, ok: bool=True): self.json = data; self.ok = ok
class BuiltinEditor:
    """
    A class that edits the builtin functions to support metadata for open function and safe closure for input functions from KeyboardInterrupt.
    """
    __slots__ = ()
    def __init__(self):
        if main_os == "Windows": return
        def holding_open(path, mode="r", *args, **kwargs):
            if "r" in mode and "+" not in mode and "w" not in mode and "a" not in mode:
                try: fd = os.open(path, os.O_RDONLY | os.O_NOATIME); return os.fdopen(fd, mode, *args, **kwargs)
                except (AttributeError, PermissionError, OSError): pass
            return _original_open(path, mode, *args, **kwargs)
        def holding_input(*args, **kwargs):
            try: return _original_input(*args, **kwargs)
            except KeyboardInterrupt: return sys.exit()
        _original_open = builtins.open
        _original_input = builtins.input
        builtins.open = holding_open
        builtins.input = holding_input
class IterableSetup:
    """
    A class that allows you to create an iterable class object to use for __iter__.
    """
    __slots__ = ("data", "index")
    def __init__(self, data: typing.Iterable): 
        self.data = data
        self.index = 0
    def __iter__(self): return self
    def __next__(self):
        if self.index < len(self.data):
            item = self.data[self.index]
            self.index += 1
            return item
        else:
            self.data = None
            self.index = 0
            raise StopIteration
class PyKitsIsAModule(Exception): pass

# Extended Namespace
requests = request
Request = request
Requests = request
Pip = pip
Plist = plist

# Throw Exception for Main
if __name__ == "__main__": raise PyKitsIsAModule("PyKits is a module and not a runable instance!")