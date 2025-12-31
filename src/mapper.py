"""
Category and client mapping.
"""

from src.config import get_category_mapping


def map_category(outlook_category: str) -> str | None:
    """Map Outlook category to SharePoint category."""
    mapping = get_category_mapping()["mapping"]
    return mapping.get(outlook_category.upper())


def extract_client_from_domains(external_domains: str, domain_to_client: dict) -> str | None:
    """Extract client name from external domains."""
    if not external_domains:
        return None

    for domain in external_domains.split(","):
        domain = domain.strip().lower()
        # Check exact match
        if domain in domain_to_client:
            return domain_to_client[domain]
        # Check if domain contains client keyword
        for known_domain, client in domain_to_client.items():
            if known_domain in domain or domain in known_domain:
                return client

    return None


def extract_client_from_title(title: str, keywords_to_client: dict) -> str | None:
    """Extract client name from event title."""
    title_lower = title.lower()

    for keyword, client in keywords_to_client.items():
        if keyword.lower() in title_lower:
            return client

    return None


def detect_client(event: dict, domain_to_client: dict, keywords_to_client: dict) -> str | None:
    """Detect client from event using domains first, then title."""
    # Priority 1: external domains
    client = extract_client_from_domains(event.get("external_domains", ""), domain_to_client)
    if client:
        return client

    # Priority 2: title keywords
    return extract_client_from_title(event.get("title", ""), keywords_to_client)
