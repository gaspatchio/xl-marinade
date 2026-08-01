# ABOUTME: Error types for mutation system validation and replay
# ABOUTME: Includes base MutationError and specific error subclasses


class MutationError(Exception):
    """Base class for mutation errors."""

    pass


class MutationValidationError(MutationError):
    """Mutation failed validation."""

    def __init__(self, mutation_id: int, message: str):
        self.mutation_id = mutation_id
        super().__init__(f"Mutation {mutation_id}: {message}")


class MutationConflictError(MutationError):
    """Mutation conflicts with current state."""

    def __init__(self, mutation_id: int, message: str):
        self.mutation_id = mutation_id
        super().__init__(f"Mutation {mutation_id}: {message}")


class MutationSequenceError(MutationError):
    """Mutation IDs are not sequential."""

    pass
