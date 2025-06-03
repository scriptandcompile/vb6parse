pub enum ParserResults<T, E>
where
    T: Clone,
    E: Clone,
{
    Success(T),
    SuccessWithFailures(T, Vec<E>),
    Failures(Vec<E>),
}

impl<T, E> ParserResults<T, E>
where
    T: Clone,
    E: Clone,
{
    pub fn is_success(&self) -> bool {
        matches!(self, ParserResults::Success(_))
    }

    pub fn is_failure(&self) -> bool {
        matches!(self, ParserResults::Failures(_))
    }

    pub fn is_success_with_failures(&self) -> bool {
        matches!(self, ParserResults::SuccessWithFailures(_, _))
    }

    pub fn has_failures(&self) -> bool {
        matches!(
            self,
            ParserResults::SuccessWithFailures(_, _) | ParserResults::Failures(_)
        )
    }

    pub fn failures(&self) -> Option<&Vec<E>> {
        match self {
            ParserResults::Success(_) => None,
            ParserResults::SuccessWithFailures(_, failures) => Some(failures),
            ParserResults::Failures(failures) => Some(failures),
        }
    }

    pub fn success(&self) -> Option<&T> {
        match self {
            ParserResults::Success(value) => Some(value),
            ParserResults::SuccessWithFailures(value, _) => Some(value),
            ParserResults::Failures(_) => None,
        }
    }

    pub fn unwrap(self) -> T {
        match self {
            ParserResults::Success(value) => value,
            ParserResults::SuccessWithFailures(value, _) => value,
            ParserResults::Failures(_) => panic!("Called unwrap on a failure result"),
        }
    }

    pub fn append_failure(&mut self, failure: E) {
        match self {
            ParserResults::Success(value) => {
                *self = ParserResults::SuccessWithFailures(value.clone(), vec![failure]);
            }
            ParserResults::SuccessWithFailures(_, failures) => {
                failures.push(failure);
            }
            ParserResults::Failures(failures) => {
                failures.push(failure);
            }
        }
    }

    pub fn extend_failures(&mut self, failures: Vec<E>) {
        match self {
            ParserResults::Success(value) => {
                *self = ParserResults::SuccessWithFailures(value.clone(), failures);
            }
            ParserResults::SuccessWithFailures(_, existing_failures) => {
                existing_failures.extend(failures);
            }
            ParserResults::Failures(existing_failures) => {
                existing_failures.extend(failures);
            }
        }
    }

    pub fn into_inner(self) -> (Option<T>, Vec<E>) {
        match self {
            ParserResults::Success(value) => (Some(value), vec![]),
            ParserResults::SuccessWithFailures(value, failures) => (Some(value), failures),
            ParserResults::Failures(failures) => (None, failures),
        }
    }
}
