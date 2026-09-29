import unittest
from unittest.mock import AsyncMock, patch

from langchain_core.messages import AIMessage
from langchain_core.language_models.fake_chat_models import FakeListChatModel

from agents import graph
from main import concise_write_answer


class FakeClient:
    async def get_tools(self):
        return []


class ProviderRetryTests(unittest.IsolatedAsyncioTestCase):
    async def run_model(self, failure, failures_before_success):
        attempts = 0

        def make_agent(**kwargs):
            retry = kwargs["middleware"][1]

            class Agent:
                async def ainvoke(self, *_args, **_kwargs):
                    async def model_call(_request):
                        nonlocal attempts
                        attempts += 1
                        if attempts <= failures_before_success:
                            raise ValueError(failure)
                        return {"messages": [AIMessage(content="Completed request")]}

                    return await retry.awrap_model_call(None, model_call)

            return Agent()

        with patch.object(graph, "get_mcp_client", return_value=FakeClient()), patch.object(
                graph, "ChatOpenRouter", return_value=object()), patch.object(
                graph, "create_agent", side_effect=make_agent), patch(
                "agents.graph.asyncio.sleep", new_callable=AsyncMock) as sleep:
            result = await graph.aask_agent("Create a sheet")
        return result, attempts, [call.args[0] for call in sleep.await_args_list]

    async def test_overloaded_response_retries_with_exponential_backoff(self):
        result, attempts, delays = await self.run_model(
            "Response validation failed: provider_overloaded", 3)
        self.assertEqual(attempts, 4)
        self.assertEqual(delays, [2.0, 4.0, 8.0])
        self.assertEqual(result["answer"], "Completed request")

    async def test_exhausted_overload_returns_one_clean_error(self):
        result, attempts, delays = await self.run_model(
            "Response validation failed: provider_overloaded", 4)
        self.assertEqual(attempts, 4)
        self.assertEqual(delays, [2.0, 4.0, 8.0])
        self.assertEqual(result["answer"],
                         "The model provider is temporarily overloaded. Please try again shortly.")
        self.assertNotIn("validation", concise_write_answer(result, [], []))
        result["operations"] = [{"name": "create_worksheet", "args": {"sheet_name": "haris"}}]
        self.assertIn("temporarily overloaded", concise_write_answer(result, [], []))
        self.assertNotIn("Created", concise_write_answer(result, [], []))

    async def test_unavailable_keeps_existing_retry_schedule(self):
        result, attempts, delays = await self.run_model("provider_unavailable", 3)
        self.assertEqual(attempts, 3)
        self.assertEqual(delays, [1.0, 2.0])
        self.assertIn("temporarily unavailable", result["answer"])
        result["operations"] = [{"name": "create_worksheet", "args": {"sheet_name": "haris"}}]
        self.assertIn("temporarily unavailable", concise_write_answer(result, [], []))

    async def test_real_agent_construction_and_model_retry(self):
        attempts = 0
        original_generate = FakeListChatModel._agenerate

        async def flaky_generate(model, *args, **kwargs):
            nonlocal attempts
            attempts += 1
            if attempts < 3:
                raise ValueError("Response validation failed: provider_overloaded")
            return await original_generate(model, *args, **kwargs)

        with patch.object(graph, "get_mcp_client", return_value=FakeClient()), patch.object(
                graph, "ChatOpenRouter", return_value=FakeListChatModel(responses=["ready"])), patch.object(
                FakeListChatModel, "_agenerate", flaky_generate), patch(
                "agents.graph.asyncio.sleep", new_callable=AsyncMock) as sleep:
            result = await graph.aask_agent("Hello", thread_id="provider_retry_agent_build")
        self.assertEqual(result["answer"], "ready")
        self.assertEqual(attempts, 3)
        self.assertEqual([call.args[0] for call in sleep.await_args_list], [2.0, 4.0])


if __name__ == "__main__":
    unittest.main()
