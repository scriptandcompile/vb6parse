pub enum ParserResults<T, E> {
    Success(T),
    SuccessWithFailures(T, Vec<E>),
    Failures(E),
}

impl<T, E> ParserResults<T, E> {
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
            ParserResults::Failures(failures) => Some(&vec![failures.clone()]),
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

    pub fn append_failures(&mut self, failures: Vec<E>) {
        match self {
            ParserResults::Success(_) => {
                *self = ParserResults::SuccessWithFailures(self.unwrap(), failures);
            }
            ParserResults::SuccessWithFailures(_, existing_failures) => {
                existing_failures.append(failures);
            }
            ParserResults::Failures(existing_failures) => {
                existing_failures.append(failures);
            }
        }
    }

    pub fn append_failure(&mut self, failure: E) {
        match self {
            ParserResults::Success(_) => {
                *self = ParserResults::SuccessWithFailures(self.unwrap(), vec![failure]);
            }
            ParserResults::SuccessWithFailures(_, existing_failures) => {
                existing_failures.push(failure);
            }
            ParserResults::Failures(existing_failures) => {
                existing_failures.push(failure);
            }
        }
    }
}
