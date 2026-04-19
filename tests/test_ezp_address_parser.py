"""
Тесты координатного парсера адресатов ЭЗП и нормализатора адреса.
Эталонные значения — из tests/fixtures/ezp/Отправка.xlsx (1-3.pdf)
и tests/fixtures/ezp/Отправка_apr.xlsx (4-5.pdf).
"""
from pathlib import Path

import pytest

from src.ezp_processor import AddressNormalizer, AddressParser, Recipient

FIXTURES = Path(__file__).parent / "fixtures" / "ezp"


@pytest.fixture(scope="module")
def parser() -> AddressParser:
    return AddressParser()


@pytest.fixture(scope="module")
def normalizer() -> AddressNormalizer:
    return AddressNormalizer()


@pytest.mark.parametrize(
    "pdf_name,expected_recipient,expected_addr_substrings,expected_type,expected_doc",
    [
        (
            "1.pdf",
            "Прокуратура города Москвы",
            ["пл. Крестьянская Застава", "д. 1", "109992", "г. Москва"],
            1,
            "01-13-1492/26",
        ),
        (
            "2.pdf",
            "ООО «НГС»",
            ["1-я Магистральная ул.", "д. 17", "стр. 4", "123007", "г. Москва"],
            1,
            "01-13-1300/26",
        ),
        (
            "3.pdf",
            "Депутату Совета депутатов муниципального округа Крылатское города Москвы Прописновой Е.О.",
            ["Осенний бул.", "д. 12", "корп. 3", "121614", "г. Москва"],
            1,
            "01-17-9/26",
        ),
        (
            "4.pdf",
            "Межрайонная природоохранная прокуратура города Москвы",
            ["ул. Профсоюзная", "д. 41", "117420", "г. Москва"],
            1,
            "01-16.5-584/26",
        ),
        (
            "5.pdf",
            "ООО «Дмитровка-Эстейт»",
            ["ул. Малая Дмитровка", "д. 12", "стр. 1", "127006", "г. Москва"],
            1,
            "01-30-2195/26",
        ),
    ],
)
def test_extract_recipients(
    parser, pdf_name, expected_recipient, expected_addr_substrings, expected_type, expected_doc
):
    pdf = FIXTURES / pdf_name
    assert pdf.exists(), f"Нет фикстуры {pdf}"

    recipients = parser.extract_recipients(str(pdf))
    assert len(recipients) == 1, f"{pdf_name}: ожидаем одного адресата, получили {len(recipients)}"

    rec = recipients[0]
    assert isinstance(rec, Recipient)
    assert rec.recipient_name == expected_recipient, (
        f"{pdf_name}: RECIPIENT\n  ожидали: {expected_recipient!r}\n  получили: {rec.recipient_name!r}"
    )
    for substr in expected_addr_substrings:
        assert substr in rec.address_raw, (
            f"{pdf_name}: адрес не содержит {substr!r}, адрес={rec.address_raw!r}"
        )
    assert rec.entity_type == expected_type, (
        f"{pdf_name}: RECIPIENT_TYPE ожидаем {expected_type}, получили {rec.entity_type}"
    )

    doc_number = parser.extract_document_number_from_pdf(str(pdf))
    assert doc_number == expected_doc, (
        f"{pdf_name}: LETTER_REG_NUMBER ожидаем {expected_doc}, получили {doc_number}"
    )


def test_extract_address_from_pdf_backward_compat(parser):
    """Старый API должен продолжать работать и возвращать непустую строку."""
    pdf = FIXTURES / "1.pdf"
    result = parser.extract_address_from_pdf(str(pdf))
    assert result is not None
    assert "Прокуратура" in result
    assert "109992" in result


def test_no_recipient_for_bogus_path(parser):
    result = parser.extract_recipients(str(FIXTURES / "does_not_exist.pdf"))
    assert result == []


@pytest.mark.parametrize(
    "raw,expected",
    [
        (
            "пл. Крестьянская Застава, д. 1,, г. Москва, 109992",
            "109992, пл. Крестьянская Застава, д. 1, г. Москва",
        ),
        (
            "1-я Магистральная ул., д. 17, стр. 4,, г. Москва, 123007",
            "123007, 1-я Магистральная ул., д. 17, стр. 4, г. Москва",
        ),
        (
            "Осенний бул., д. 12, корп. 3,, г. Москва, 121614",
            "121614, Осенний бул., д. 12, корп. 3, г. Москва",
        ),
        (
            "ул. Профсоюзная, д. 41,, г. Москва, 117420",
            "117420, ул. Профсоюзная, д. 41, г. Москва",
        ),
        (
            "ул. Малая Дмитровка, д. 12 стр. 1,, г. Москва, 127006",
            "127006, ул. Малая Дмитровка, д. 12 стр. 1, г. Москва",
        ),
    ],
)
def test_normalize_address(normalizer, raw, expected):
    assert normalizer.normalize_address(raw) == expected


def test_normalizer_handles_empty():
    assert AddressNormalizer().normalize_address("") == ""


def test_normalizer_no_double_periods_on_abbreviated():
    """Регрессия: старый нормалайзер превращал 'ул.' в 'ул..' (баг word-boundary)."""
    result = AddressNormalizer().normalize_address("ул. Тверская, д. 1, г. Москва, 125009")
    assert "ул.." not in result
    assert "д.." not in result
    assert result == "125009, ул. Тверская, д. 1, г. Москва"


def test_normalizer_expands_full_forms():
    """Если в адресе полные формы ('улица', 'дом') — приводим к сокращениям."""
    result = AddressNormalizer().normalize_address("улица Тверская, дом 1, г. Москва, 125009")
    assert result == "125009, ул. Тверская, д. 1, г. Москва"


def test_email_is_captured_not_joined_to_address(parser):
    """
    Email в блоке адресата не должен попадать ни в recipient_name, ни в address_raw.
    PDF 1.pdf и 3.pdf содержат email-строки.
    """
    for pdf_name in ("1.pdf", "3.pdf"):
        recipients = parser.extract_recipients(str(FIXTURES / pdf_name))
        rec = recipients[0]
        assert "@" not in rec.recipient_name, f"{pdf_name}: email просочился в имя"
        assert "@" not in rec.address_raw, f"{pdf_name}: email просочился в адрес"
        assert rec.emails, f"{pdf_name}: email должен быть отдельно зафиксирован"
