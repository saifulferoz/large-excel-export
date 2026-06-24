<?php

namespace App\Service;

class DataGeneratorService
{
    public function generate(int $count): \Generator
    {
        $categories = ['Electronics', 'Furniture', 'Clothing', 'Books', 'Toys', 'Office Supplies'];
        $baseTime = time();

        $dates = [];
        for ($j = 0; $j < 100; $j++) {
            $dates[$j] = date('d-m-Y', $baseTime - ($j * 3600));
        }

        for ($i = 1; $i <= $count; ++$i) {
            // Keep memory footprint low by avoiding complex date/string manipulation
            yield [
                'id' => $i,
                'name' => 'Product_' . $i,
                'category' => $categories[$i % 6],
                'quantity' => ($i % 99) + 1,
                'price' => (float) (($i * 13) % 490 + 10.5),
                'date' => $dates[$i % 100],
            ];
        }
    }
}
