<?php

return (new PhpCsFixer\Config())
    ->setRules([
        'native_function_invocation' => false,
    ])
    ->setFinder(
        PhpCsFixer\Finder::create()
            ->in(__DIR__)
            ->append([__FILE__])
    )
;
