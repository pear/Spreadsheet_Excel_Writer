<?php

return (new PhpCsFixer\Config())
    ->setRules([
        '@PER-CS' => true,
        'native_function_invocation' => false,
        'array_syntax' => false,
        'concat_space' => false,
        'blank_line_after_opening_tag' => false,
        'visibility_required' => false,
        'trailing_comma_in_multiline' => false,
        'method_argument_space' => false,
        'array_indentation' => false,
        'single_space_around_construct' => false,
        'braces_position' => false,
        'statement_indentation' => false,
        'binary_operator_spaces' => false,
        'single_blank_line_at_eof' => false,
        'elseif' => false,
        'control_structure_braces' => false,
        'control_structure_continuation_position' => false,
        'no_multiple_statements_per_line' => false,
    ])
    ->setFinder(
        PhpCsFixer\Finder::create()
            ->in(__DIR__)
            ->append([__FILE__])
    )
;
