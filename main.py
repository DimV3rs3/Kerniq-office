"""Office plugin. The tools live in tools/ and are offered to the model."""

from __future__ import annotations

from Kerniq.plugins.base import Plugin, PluginContext
from Kerniq.plugins.manifest import PluginManifest


class PluginImpl(Plugin):
    def __init__(self) -> None:
        self._manifest = PluginManifest(
            plugin_id="office",
            name="Office",
            version="0.1.1",
            api_version="1.0",
            manifest_version="1.0",
            entry_point="main.PluginImpl",
            description=(
                "Creates and edits Microsoft Office and LibreOffice documents: "
                "numbered tables, fonts, and styles. Not plain text."
            ),
            capabilities=["tools"],
            author="Kerniq",
            license="MIT",
        )

    @property
    def manifest(self) -> PluginManifest:
        return self._manifest

    async def initialize(self, context: PluginContext) -> bool:
        self._context = context
        return True

    async def shutdown(self) -> None:
        return None

    def settings_schema(self):
        from Kerniq.interfaces.settings_schema import SettingsSchema, SettingField, SettingType
        return SettingsSchema(
            component_id="office",
            component_type="plugin",
            title="Office",
            fields=[
                SettingField(key="enabled", label="Enabled", type=SettingType.BOOL, default=True),
            ],
        )

    def apply_settings(self, values): pass
