import re

ip_pattern = re.compile(
    r"(25[0-5]|2[0-4]\d|1\d{2}|[1-9]?\d)" r"(\.(25[0-5]|2[0-4]\d|1\d{2}|[1-9]?\d)){3}"
)


def extract_ip(text):
    match = ip_pattern.search(text)
    return match.group() if match else None


def test_single_ip():
    assert extract_ip("IP-172.16.2.54") == "172.16.2.54"


def test_no_ip():
    assert extract_ip("No IP here") is None


def test_embedded_ip():
    text = "Device (IP-192.168.1.1, Name-X)"
    assert extract_ip(text) == "192.168.1.1"
