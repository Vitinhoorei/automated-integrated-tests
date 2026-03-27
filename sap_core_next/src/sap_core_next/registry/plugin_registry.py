from __future__ import annotations

from dataclasses import dataclass, field
from typing import List

from sap_core_next.ports.plugin import ModulePlugin


@dataclass
class PluginRegistry:
    plugins: List[ModulePlugin] = field(default_factory=list)

    def register(self, plugin: ModulePlugin) -> None:
        self.plugins.append(plugin)

    def resolve(self, module: str) -> ModulePlugin:
        for plugin in self.plugins:
            if plugin.supports(module):
                return plugin
        raise KeyError(f"No plugin registered for module '{module}'")
