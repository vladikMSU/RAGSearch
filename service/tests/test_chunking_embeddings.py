import unittest

from ragsearch_service.chunking import chunk_text
from ragsearch_service.embeddings import HashingEmbedder, cosine_for_normalized


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


if __name__ == "__main__":
    unittest.main()

