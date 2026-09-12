from dataclasses import dataclass


@dataclass(frozen=True)
class RunConfig:
    interval: float = 1.0
    input_mode: str = "left"
    keys: tuple[int, ...] = ()
    key_label: str = "SPACE"
    saved_point: tuple[int, int] | None = None
    max_actions: int | None = None
    max_duration: float | None = None
    delay: float = 0.0
    unity_only: bool = False
    preview: bool = False


class Session:
    def __init__(self, native):
        self.native = native
        self.config = RunConfig()
        self.active = False
        self.count = 0
        self.elapsed = 0.0
        self.started_at = None
        self.activate_at = 0.0
        self.next_action_at = 0.0
        self.status = "Idle"
        self.events = []

    def start(self, config: RunConfig, now: float) -> None:
        self.config = config
        self.active = True
        self.count = 0
        self.elapsed = 0.0
        self.started_at = None
        self.activate_at = now + config.delay
        self.next_action_at = self.activate_at
        self.status = "Waiting for start delay" if config.delay else "Ready"
        self.events.clear()

    def stop(self, message: str, now: float) -> None:
        if self.active and self.started_at is not None:
            self.elapsed = max(0.0, now - self.started_at)
        self.active = False
        self.status = message

    def tick(self, now: float) -> None:
        if not self.active:
            return
        config = self.config
        if now < self.activate_at:
            self.status = f"Starting in {self.activate_at - now:.1f}s; switch to your target"
            return
        if self.started_at is None:
            self.started_at = now
        self.elapsed = max(0.0, now - self.started_at)
        if config.max_duration is not None and self.elapsed >= config.max_duration:
            self.stop("Completed duration limit", now)
            return
        if config.unity_only and not config.preview:
            if self.native.foreground_executable() != "unity.exe":
                self.stop("Stopped: Unity Editor is not focused", now)
                return
        if now < self.next_action_at:
            return
        if not config.preview and self.native.keys_held(config.keys):
            self.status = "Waiting for you to release keys"
            return
        if config.unity_only and not config.preview and config.input_mode != "keyboard":
            if not self.native.mouse_target_is_unity(config.saved_point):
                self.stop("Stopped: the mouse target is outside Unity Editor", now)
                return
        action = f"key {config.key_label}" if config.input_mode == "keyboard" else f"{config.input_mode} click"
        if config.saved_point is not None and config.input_mode != "keyboard":
            action += f" at {config.saved_point}"
        try:
            if not config.preview:
                self.native.perform(config.input_mode, config.keys, config.saved_point)
        except OSError as error:
            self.stop(str(error), now)
            return
        self.count += 1
        self.events.append(f"{'Preview' if config.preview else 'Sent'} #{self.count}: {action}")
        self.status = f"{'Previewing' if config.preview else 'Active'}: {action}"
        self.next_action_at = now + config.interval
        if config.max_actions is not None and self.count >= config.max_actions:
            self.stop("Completed count limit", now)
