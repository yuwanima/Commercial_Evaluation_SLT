
from google import genai
from google.genai import types

import rag_config


def prepare_document_input(text, title=None):
    """
    Format a procurement retrieval unit for Gemini Embedding 2.
    """
    title = title or "none"

    return f"title: {title} | text: {text}"


def prepare_query_input(query):
    """
    Format a user's question for procurement document retrieval.
    """
    return f"task: question answering | query: {query}"


def create_vertex_client():
    """
    Create a Vertex AI client.
    This function does not generate embeddings.
    """
    rag_config.validate_vertex_settings()

    return genai.Client(
        vertexai=True,
        project=rag_config.GOOGLE_CLOUD_PROJECT,
        location=rag_config.GOOGLE_CLOUD_LOCATION
    )


def embed_text(client, text):
    """
    Generate one embedding when explicitly called.
    """
    response = client.models.embed_content(
        model=rag_config.EMBEDDING_MODEL,
        contents=text,
        config=types.EmbedContentConfig(
            output_dimensionality=rag_config.EMBEDDING_DIMENSIONS
        )
    )

    if not response.embeddings:
        raise RuntimeError("Google returned no embeddings.")

    values = response.embeddings[0].values

    if len(values) != rag_config.EMBEDDING_DIMENSIONS:
        raise RuntimeError(
            f"Expected {rag_config.EMBEDDING_DIMENSIONS} dimensions, "
            f"received {len(values)}."
        )

    return list(values)
