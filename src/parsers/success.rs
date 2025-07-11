pub enum Success<T, E> {
    Value(T),
    ValueWithFailures(T, Vec<E>),
}

impl<T, E> Success<T, E>
where
    T: Clone,
{
    pub fn value(&self) -> &T {
        match self {
            Success::Value(value) => value,
            Success::ValueWithFailures(value, _) => value,
        }
    }

    pub fn failures(&self) -> Option<&Vec<E>> {
        match self {
            Success::Value(_) => None,
            Success::ValueWithFailures(_, failures) => Some(failures),
        }
    }

    pub fn has_failures(&self) -> bool {
        matches!(self, Success::ValueWithFailures(_, _))
    }

    pub fn unwrap(self) -> T {
        match self {
            Success::Value(value) => value,
            Success::ValueWithFailures(value, _) => value,
        }
    }

    pub fn append_failure(&mut self, failure: E) {
        match self {
            Success::Value(value) => {
                *self = Success::ValueWithFailures(value.clone(), vec![failure]);
            }
            Success::ValueWithFailures(_, failures) => {
                failures.push(failure);
            }
        }
    }
}
