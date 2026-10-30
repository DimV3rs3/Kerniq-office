"""Office follows the folder-tool plugin standard and can read the sample files."""

from __future__ import annotations

import asyncio
import importlib.util
from pathlib import Path

from Kerniq.plugins.base import PluginContext
from Kerniq.plugins.manager import PluginManager
from Kerniq.tools.registry import ToolRegistry


def _tools_module():
    path = Path(__file__).resolve().parents[1] / "tools" / "__init__.py"
    spec = importlib.util.spec_from_file_location("kerniq_office_tools_test", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_office_discovers_and_exposes_its_tools():
    plugins_root = Path(__file__).resolve().parents[2]
    registry = ToolRegistry()
    manager = PluginManager(
        lambda: PluginContext(logger=None, config=None, event_bus=None, tool_registry=registry)
    )

    async def run():
        await manager.discover_and_load(plugins_root)
        return manager.plugins.get("office")

    plugin = asyncio.run(run())
    assert plugin is not None
    assert manager.manifests["office"].plugin_id == "office"
    names = {tool.name for tool in registry.agent_tools()}
    assert {"office_inspect", "office_read", "office_create", "office_edit"} <= names


def test_office_reads_the_sample_documents_through_virtual_paths():
    module = _tools_module()
    workspace = Path(__file__).resolve().parents[3] / "workspace"

    root = str(workspace)

    class Context:
        pass

    Context.workspace = root

    xlsx = module._safe("./workspace/Отчёт_проверки_ФОС (14).xlsx", Context(), must_exist=True)
    docx = module._safe(
        "./workspace/ФОС_2025_09.03.03_итцэ_О_Б1.О.09_ГиСД_Подобай.docx",
        Context(),
        must_exist=True,
    )
    engine = module._engine()
    book = engine.read_document(xlsx)
    assert book.get("sheets")
    document = engine.read_document(docx)
    assert document.get("paragraphs") or document.get("tables")
    info = engine.inspect_document(xlsx)
    assert "words" in info
