<?php

return (new PhpCsFixer\Config())
    ->setRules([
        '@PER-CS' => true,
        'array_syntax' => false,
        'concat_space' => false,
        'visibility_required' => ['elements' => ['property', 'method']], // Exclude 'const' for PHP 5.6 compatibility
        'trailing_comma_in_multiline' => ['elements' => ['arrays']],
        'method_argument_space' => false,
        'array_indentation' => false,
        'braces_position' => false,
    ])
    ->setFinder(
        PhpCsFixer\Finder::create()
            ->in(__DIR__)
            ->append([__FILE__])
    )
;
