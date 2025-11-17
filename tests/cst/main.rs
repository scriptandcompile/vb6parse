// This whole file is just a horrible hack to allow grouping of the CST parsing tests into
// a separate folder without needing to throw all the tests into a single huge file or
// throwing all test files in the root tests folder.

mod call;
mod comments;
mod for_loop;
mod do_loop;
mod set_statement;
mod conditional;
mod cst_navigation;
mod cst;
mod declaration;
mod option_explicit;
mod sub;
mod assignment;
mod label;
mod with_statement;
mod select_case;
mod goto_statement;
mod inline_if;
mod exit_statement;
mod property_statement;
mod appactivate_statement;
mod beep_statement;
mod chdir_statement;
