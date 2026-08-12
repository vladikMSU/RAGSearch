import tempfile
import unittest
from pathlib import Path

from ragsearch_service.chunking import chunk_text
from ragsearch_service.embeddings import (
    HashingEmbedder,
    _fingerprint_model_directory,
    cosine_for_normalized,
    create_embedder,
)


class ChunkingTests(unittest.TestCase):
    def test_chunks_are_bounded_and_deterministic(self) -> None:
        text = " ".join(f"word-{index}" for index in range(300))
        first = list(chunk_text(text, max_chars=128, overlap_chars=24))
        second = list(chunk_text(text, max_chars=128, overlap_chars=24))
        self.assertEqual(first, second)
        self.assertGreater(len(first), 2)
        self.assertTrue(all(0 < len(item) <= 128 for item in first))

    def test_empty_text_creates_no_chunks(self) -> None:
        self.assertEqual([], list(chunk_text(" \r\n ")))


class HashingEmbeddingTests(unittest.TestCase):
    def test_embedding_is_deterministic_and_normalized(self) -> None:
        embedder = HashingEmbedder(dimensions=64)
        first = embedder.embed("Semantic Outlook search")
        second = embedder.embed("Semantic Outlook search")
        unrelated = embedder.embed("banana telescope")
        self.assertEqual(first, second)
        self.assertAlmostEqual(cosine_for_normalized(first, first), 1.0, places=6)
        self.assertGreater(
            cosine_for_normalized(first, second),
            cosine_for_normalized(first, unrelated),
        )
        self.assertEqual(embedder.fingerprint, HashingEmbedder(dimensions=64).fingerprint)
        self.assertNotEqual(embedder.fingerprint, HashingEmbedder(dimensions=32).fingerprint)

    def test_model_artifact_fingerprint_changes_with_file_content(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            artifact = root / "model.bin"
            artifact.write_bytes(b"first model weights")
            first = _fingerprint_model_directory(root)
            artifact.write_bytes(b"second model weights")
            second = _fingerprint_model_directory(root)

        self.assertRegex(first, r"\Asha256:[0-9a-f]{64}\Z")
        self.assertNotEqual(first, second)

    def test_factory_rejects_noncanonical_provider_names(self) -> None:
        self.assertIsInstance(create_embedder("hash"), HashingEmbedder)
        with self.assertRaisesRegex(ValueError, "only valid"):
            create_embedder("hash", "ignored-model")
        for provider in (
            "HASH",
            "hashing",
            "fallback",
            " hash ",
            "sentence_transformers",
            "st",
        ):
            with self.subTest(provider=provider):
                with self.assertRaisesRegex(ValueError, "Unknown embedding provider"):
                    create_embedder(provider)


if __name__ == "__main__":
    unittest.main()
