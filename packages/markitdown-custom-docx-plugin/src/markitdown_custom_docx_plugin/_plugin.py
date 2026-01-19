import locale
from mammoth import convert_to_html
from mammoth.results import Result
from typing import BinaryIO, Any
from bs4 import BeautifulSoup
from bs4.element import Tag
import base64


from markitdown import (
    MarkItDown,
    DocumentConverter,
    DocumentConverterResult,
    StreamInfo,
)


__plugin_interface_version__ = (
    1  # The version of the plugin interface that this plugin uses
)

ACCEPTED_MIME_TYPE_PREFIXES = [
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
]

ACCEPTED_FILE_EXTENSIONS = [".docx"]


def register_converters(markitdown: MarkItDown, **kwargs):
    """
    Called during construction of MarkItDown instances to register converters provided by plugins.
    """

    # Simply create and attach an CustiomDocxConverter instance
    markitdown.register_converter(CustomDocxConverter())


class CustomDocxConverter(DocumentConverter):
    """
    Converts a DOCX file.
    """
    def __convert_to_html(
        self,
        stream_data: str,
        **kwargs: Any
    ) -> Result:
        return convert_to_html(stream_data)

    def __convert_to_markdown(
        self,
        html: str,
        **kwargs: Any
    ) -> str:
        markdown = []
        soup = BeautifulSoup(html, "html.parser")
        for element in soup.body.descendants:
            if isinstance(element, Tag):
                if element.name == "span":
                    text = element.get_text(strip=True)
                    markdown.append(text)
                elif element.name == "table":
                    table_content = str(element)
                    markdown.append(self.__process_table(table_content))
                elif element.name == "img" and element.get("src"):
                    src = element["src"]
                    if ".fld/" in src:
                        markdown.append(self.__process_image(src))
                elif element.name == "a" and element.get("href"):
                    href = element["href"]
                    extension = href.split(".")[-1].lower()
                    if ".fld/" in href and extension.endswith((".png", ".jpg", ".jpeg", ".gif")):
                        markdown.append(self.__process_image(href, extension))
        return "\n".join(markdown)

    def __process_table(
        self,
        table_content: str,
    ) -> str:
        pass

    def __process_image(
        self,
        image_path: str,
        image_extension: str,
    ) -> str:
        with open(image_path, "rb") as img_file:
            img_data = img_file.read()
            encoded_img = base64.b64encode(img_data).decode("utf-8")
            return f"![Image](data:image/{image_extension};base64,{encoded_img})"

    def accepts(
        self,
        file_stream: BinaryIO,
        stream_info: StreamInfo,
        **kwargs: Any,
    ) -> bool:
        mimetype = (stream_info.mimetype or "").lower()
        extension = (stream_info.extension or "").lower()

        if extension in ACCEPTED_FILE_EXTENSIONS:
            return True

        for prefix in ACCEPTED_MIME_TYPE_PREFIXES:
            if mimetype.startswith(prefix):
                return True

        return False

    def convert(
        self,
        file_stream: BinaryIO,
        stream_info: StreamInfo,
        **kwargs: Any,
    ) -> DocumentConverterResult:
        # Read the file stream into an str using hte provided charset encoding, or using the system default
        encoding = stream_info.charset or locale.getpreferredencoding()
        stream_data = file_stream.read().decode(encoding)
        html_data = self.__convert_to_html(stream_data).value
        markdown = self.__convert_to_markdown(html_data)
        # Return the result
        return DocumentConverterResult(
            title=None,
            markdown=markdown,
        )
